using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace MPHBSMS
{
    // ==========================================================
    // 1. HELPER DEFINITIONS (Enums, Request Models, Result)
    // ==========================================================
    public enum ConversationState
    {
        Greeting,
        AskForPrimaryAction,
        AskForStartDate,
        AskForStartMonth, // NEW: Sequential date state
        AskForStartYear,  // NEW: Sequential date state
        AskForEndDate,
        AskForEndMonth,   // NEW: Sequential date state
        AskForEndYear,    // NEW: Sequential date state
        AskForDataScope,
        AskForDataAction,
        AskForSaveFormat,
        AskForReportToOpen,
        ConfirmationAndGenerate,
        Complete
    }

    public class ReportRequest
    {
        private List<string> _categories;
        public string PrimaryAction { get; set; }
        public DateTime StartDate
        {
            get;
            set;
        }
        public DateTime EndDate
        {
            get;
            set;
        }
        public string DataScope
        {
            get;
            set;
        }
        public string DataAction
        {
            get;
            set;
        }
        public string FileFormat
        {
            get;
            set;
        }
        public string ReportName
        {
            get;
            set;
        }

        // NEW: Temporary storage for separate date inputs
        public int? TempDay { get; set; }
        public int? TempMonth { get; set; }
        public int? TempYear { get; set; }

        public List<string> Categories
        {
            get
            {
                return _categories;
            }
            set
            {
                _categories = value;
            }
        }

        public ReportRequest()
        {
            _categories = new List<string>();
        }

        // NEW: Clears temporary date components
        public void ClearTempDate()
        {
            TempDay = null;
            TempMonth = null;
            TempYear = null;
        }
    }

    public class ProcessResult
    {
        public bool Success
        {
            get;
            set;
        }
        public string ErrorMessage
        {
            get;
            set;
        }
    }

    // ==========================================================
    // 2. REPORT MANAGER
    // ==========================================================
    public class ReportManager
    {
        // --- Fields & Properties ---
        private ConversationState _currentState;

        public ConversationState CurrentState
        {
            get
            {
                return _currentState;
            }
            set
            {
                _currentState = value;
            }
        }

        public ReportRequest CurrentRequest
        {
            get;
            private set;
        }

        private const string AppFolderName = "MyReportGeneratorFiles";
        private readonly string ReportDirectoryPath;
        // **********************************************
        // ACCEPTED DATE FORMATS
        // **********************************************
        private static readonly string[] AcceptedDateFormats = new string[]
        {
            // YYYY-MM-DD (ISO-like)
            "yyyy-MM-dd", 
            "yyyy/MM/dd",
            "yyyy-M-d",   
            "yyyy/M/d", 
            "yyyy-MM-dd HH:mm:ss", 
            
            // MM/DD/YYYY (US format)
            "MM-dd-yyyy", 
            "MM/dd/yyyy",
            "M-d-yyyy",   
            "M/d/yyyy", // M/d/yyyy
            "MM/dd/yyyy HH:mm:ss", 

            // DD/MM/YYYY (European format)
            "dd-MM-yyyy", 
            "dd/MM/yyyy",
            "d-M-yyyy",   
            "d/M/yyyy",
            "dd/MM/yyyy HH:mm:ss", 
        };
        // **********************************************

        private readonly List<string> ValidReportCategories = new List<string>
        {
            "Admission", "Deaths", "Transfers", "Discharges", "Consultations"
        };
        // --- Constructor ---
        public ReportManager()
        {
            CurrentRequest = new ReportRequest();
            _currentState = ConversationState.AskForPrimaryAction;
            ReportDirectoryPath = InitializeReportDirectory();
        }

        private string InitializeReportDirectory()
        {
            try
            {
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string finalPath = Path.Combine(documentsPath, AppFolderName);

                if (!Directory.Exists(finalPath))
                {
                    Directory.CreateDirectory(finalPath);
                }
                return finalPath;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // --- State machine driver ---
        public ProcessResult ProcessAndAdvance(string input)
        {
            string errorMessage = null;
            if (input == null) input = string.Empty;

            // sanitize input
            string cleanedInput = new string(input.Where(c => !char.IsControl(c)).ToArray()).Trim();
            bool isValid = ProcessInput(cleanedInput, out errorMessage);

            if (isValid)
            {
                if (CurrentState == ConversationState.ConfirmationAndGenerate)
                {
                    CurrentState = ConversationState.Complete;
                }
                else
                {
                    // Full state advance only happens if ProcessInput didn't use the special signal
                    MoveToNextState();
                }
            }
            else if (errorMessage == "NEXT_STEP_OK")
            {
                // FIX: If the input was accepted and manually advanced the state 
                // (e.g., Day -> Month), we signal success to the caller but skip MoveToNextState.
                return new ProcessResult { Success = true, ErrorMessage = null };
            }

            return new ProcessResult { Success = isValid, ErrorMessage = errorMessage };
        }

        private bool ProcessInput(string cleanedInput, out string errorMessage)
        {
            errorMessage = null;
            bool isValid = false;
            DateTime parsedDate;
            List<string> selectedCategories;
            string reportName;
            int parsedNumber;

            switch (CurrentState)
            {
                case ConversationState.Greeting:
                case ConversationState.AskForPrimaryAction:
                    string lowerInput = cleanedInput.ToLower();
                    if (lowerInput.Contains("generate") || lowerInput.Contains("create") || lowerInput.Contains("make report"))
                    {
                        CurrentRequest.PrimaryAction = "GenerateReport";
                        isValid = true;
                    }
                    else if (lowerInput.Contains("open") || lowerInput.Contains("show"))
                    {
                        CurrentRequest.PrimaryAction = "OpenExisting";
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please enter 'Generate' or 'Open'.";
                    }
                    break;

                case ConversationState.AskForStartDate:
                    CurrentRequest.ClearTempDate(); // Start of StartDate process

                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    {
                        // 1. Single date input (e.g., "2023-01-01")
                        CurrentRequest.StartDate = parsedDate;
                        isValid = true;
                    }
                    else if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 31)
                    {
                        // 2. Separate date input (e.g., "1" for day)
                        CurrentRequest.TempDay = parsedNumber;
                        // Manual state transition for sequential input
                        CurrentState = ConversationState.AskForStartMonth;
                        // Signal manual advance success without triggering MoveToNextState
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        errorMessage = "Invalid input. Please enter a full date (e.g., 2023-01-01) or just the **Day** (1-31) to start sequential input.";
                    }

                    break;

                case ConversationState.AskForStartMonth:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 12)
                    {
                        CurrentRequest.TempMonth = parsedNumber;
                        // Manual state transition for sequential input
                        CurrentState = ConversationState.AskForStartYear;
                        // Signal manual advance success without triggering MoveToNextState
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        errorMessage = "Invalid month. Please enter a number between 1 and 12.";
                    }
                    break;

                case ConversationState.AskForStartYear:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1900 && parsedNumber <= DateTime.Now.Year + 5)
                    {
                        // Use DateTime constructor
                        int year = parsedNumber;
                        // Use 1 as default if TempMonth/TempDay somehow got wiped (safety)
                        int month = CurrentRequest.TempMonth ?? 1;
                        int day = CurrentRequest.TempDay ?? 1;

                        try
                        {
                            CurrentRequest.StartDate = new DateTime(year, month, day);
                            CurrentRequest.ClearTempDate();
                            isValid = true;
                            // Will trigger MoveToNextState
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            errorMessage = "The date entered (" + day + "/" + month + "/" + year + ") is invalid (e.g., February 30th).";
                        }
                    }
                    else
                    {
                        errorMessage = "Invalid year. Please enter a four-digit year (e.g., 2024).";
                    }
                    break;

                case ConversationState.AskForEndDate:
                    CurrentRequest.ClearTempDate(); // Start of EndDate process

                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    {
                        // 1. Single date input
                        if (parsedDate < CurrentRequest.StartDate)
                        {
                            errorMessage = "The end date cannot be before the start date (" + CurrentRequest.StartDate.ToShortDateString() + ").";
                        }
                        else
                        {
                            CurrentRequest.EndDate = parsedDate;
                            isValid = true;
                        }
                    }
                    else if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 31)
                    {
                        // 2. Separate date input (e.g., "31" for day)
                        CurrentRequest.TempDay = parsedNumber;
                        // Manual state transition for sequential input
                        CurrentState = ConversationState.AskForEndMonth;
                        // Signal manual advance success
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(errorMessage))
                            errorMessage = "Unrecognized date. Please enter a full date (e.g., 2023-10-31 or 31/10/2023) or just the **Day** (1-31).";
                    }


                    break;

                case ConversationState.AskForEndMonth:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 12)
                    {
                        CurrentRequest.TempMonth = parsedNumber;
                        // Manual state transition for sequential input
                        CurrentState = ConversationState.AskForEndYear;
                        // Signal manual advance success
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        errorMessage = "Invalid month. Please enter a number between 1 and 12.";
                    }
                    break;

                case ConversationState.AskForEndYear:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1900 && parsedNumber <= DateTime.Now.Year + 5)
                    {
                        // Use DateTime constructor
                        int year = parsedNumber;
                        int month = CurrentRequest.TempMonth ?? 1;
                        int day = CurrentRequest.TempDay ?? 1;

                        try
                        {
                            parsedDate = new DateTime(year, month, day);

                            if (parsedDate < CurrentRequest.StartDate)
                            {
                                errorMessage = "The final end date (" + day + "/" + month + "/" + year + ") cannot be before the start date (" + CurrentRequest.StartDate.ToShortDateString() + ").";
                            }
                            else
                            {
                                CurrentRequest.EndDate = parsedDate;
                                CurrentRequest.ClearTempDate();
                                isValid = true;
                                // State will now advance to AskForDataScope via MoveToNextState
                            }
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            errorMessage = "The final end date entered (" + day + "/" + month + "/" + year + ") is invalid (e.g., February 30th).";
                        }
                    }
                    else
                    {
                        errorMessage = "Invalid year. Please enter a four-digit year (e.g., 2024).";
                    }
                    break;

                case ConversationState.AskForDataScope:
                    if (cleanedInput.Equals("All Data", StringComparison.OrdinalIgnoreCase) ||
                        cleanedInput.Equals("All", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.DataScope = "All Data";
                        CurrentRequest.Categories.Clear();
                        isValid = true;
                    }
                    else
                    {
                        if (TryParseCategories(cleanedInput, out selectedCategories, out errorMessage))
                        {
                            CurrentRequest.DataScope = "By Category";
                            CurrentRequest.Categories = selectedCategories;
                            isValid = true;
                        }
                    }
                    break;
                case ConversationState.AskForDataAction:
                    if (cleanedInput.Equals("Save File", StringComparison.OrdinalIgnoreCase) ||
                        cleanedInput.Equals("Save", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.DataAction = "Save File";
                        isValid = true;
                    }
                    else if (cleanedInput.Equals("Print", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.DataAction = "Print";
                        isValid = true;
                    }
                    else if (cleanedInput.Equals("Display", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.DataAction = "Display";
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please enter 'Save File', 'Print', or 'Display'.";
                    }
                    break;
                case ConversationState.AskForSaveFormat:
                    if (cleanedInput.Equals("PDF", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.FileFormat = "PDF";
                        isValid = true;
                    }
                    else if (cleanedInput.Equals("DOCX", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentRequest.FileFormat = "DOCX";
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please enter 'PDF' or 'DOCX'.";
                    }
                    break;
                case ConversationState.AskForReportToOpen:
                    if (FileExistsInReportDir(cleanedInput, out reportName, out errorMessage))
                    {
                        CurrentRequest.ReportName = reportName;
                        isValid = true;
                    }
                    break;
                case ConversationState.ConfirmationAndGenerate:
                    if (cleanedInput.Equals("Confirm", StringComparison.OrdinalIgnoreCase) ||
                        cleanedInput.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                        cleanedInput.Equals("Go", StringComparison.OrdinalIgnoreCase))
                    {
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please type 'Confirm' to proceed with generation.";
                    }
                    break;
            }

            return isValid;
        }

        private void MoveToNextState()
        {
            switch (CurrentState)
            {
                case ConversationState.Greeting:
                case ConversationState.AskForPrimaryAction:
                    CurrentState = (CurrentRequest.PrimaryAction
                        == "GenerateReport")
                        ?
                        ConversationState.AskForStartDate
                        : ConversationState.AskForReportToOpen;
                    break;

                // --- Start Date Transitions (Major Step Completion) ---
                case ConversationState.AskForStartDate:
                    // Only reached if a FULL date was entered.
                    CurrentState = ConversationState.AskForEndDate;
                    break;

                case ConversationState.AskForStartYear:
                    // Sequential date was completed.
                    CurrentState = ConversationState.AskForEndDate;
                    break;

                // --- End Date Transitions (Major Step Completion) ---
                case ConversationState.AskForEndDate:
                    // Only reached if a FULL date was entered.
                    CurrentState = ConversationState.AskForDataScope;
                    break;

                case ConversationState.AskForEndYear:
                    // Sequential date was completed.
                    CurrentState = ConversationState.AskForDataScope;
                    break;

                case ConversationState.AskForDataScope:
                    CurrentState = ConversationState.AskForDataAction;
                    break;
                case ConversationState.AskForDataAction:
                    CurrentState = (CurrentRequest.DataAction == "Save File")
                        ?
                        ConversationState.AskForSaveFormat
                        : ConversationState.ConfirmationAndGenerate;
                    break;
                case ConversationState.AskForSaveFormat:
                    CurrentState = ConversationState.ConfirmationAndGenerate;
                    break;
                case ConversationState.AskForReportToOpen:
                    CurrentState = ConversationState.ConfirmationAndGenerate;
                    break;
            }
        }

        public string GetCurrentPrompt()
        {
            switch (CurrentState)
            {
                case ConversationState.Greeting:
                    return "Welcome to the Report Generator! I'm here to guide you"
+ "\n" +
                        "through creating or opening a report.";
                case ConversationState.AskForPrimaryAction:
                    return "Do you want to **Generate** or **Open** an existing report?";

                // Updated Prompts for Hybrid/Sequential Date Input
                case ConversationState.AskForStartDate:
                    return "What is the **Start Date** for the report? You can enter the full date (e.g., '**2023-01-01**' or '**01/01/2023**') OR just the **Day** number (1-31) to start sequential input.";
                case ConversationState.AskForStartMonth:
                    return "Please enter the **Month** number for the start date (1-12).";
                case ConversationState.AskForStartYear:
                    return "Please enter the **Year** for the start date (e.g., 2024).";
                case ConversationState.AskForEndDate:
                    return "What is the **End Date** for the report? You can enter the full date (e.g., '**2023-10-31**' or '**10/31/2023**') OR just the **Day** number (1-31) to start sequential input.";
                case ConversationState.AskForEndMonth:
                    return "Please enter the **Month** number for the end date (1-12).";
                case ConversationState.AskForEndYear:
                    return "Please enter the **Year** for the end date (e.g., 2024).";
                // End Updated Prompts

                case ConversationState.AskForDataScope:
                    return "Do you want **All Data** or specific categories? (Valid: " + string.Join(", ", ValidReportCategories) + " or 'All')";
                case ConversationState.AskForDataAction:
                    return "What do you want to do with the data? (**Save File**, **Print**, or **Display**)";
                case ConversationState.AskForSaveFormat:
                    return "What format do you want to save as? (**PDF** or **DOCX**)";
                case ConversationState.AskForReportToOpen:
                    string list = string.Join("\n - ", GetAvailableReports());
                    return "What is the name of the existing report you want to open?\n\nAvailable Reports:\n - " + list;
                case ConversationState.ConfirmationAndGenerate:
                    return "Ready to proceed with " + CurrentRequest.PrimaryAction + ". Please confirm. (Type 'Confirm')";
                case ConversationState.Complete:
                    return "Operation complete. Type Generate or Open to start a new report.";
                default:
                    return "Error: Unknown state.";
            }
        }

        // --- Execution and Helpers ---
        public string ExecuteReportLogic()
        {
            if (CurrentRequest.PrimaryAction == "OpenExisting")
            {
                return "Action Complete: Attempting to open existing report file: " + CurrentRequest.ReportName;
            }

            StringBuilder actionMessage = new StringBuilder();
            actionMessage.AppendLine("Report Generation for scope: " + CurrentRequest.DataScope + " (Dates: " +
                                     CurrentRequest.StartDate.ToShortDateString() + " to " + CurrentRequest.EndDate.ToShortDateString() + ")");
            if (CurrentRequest.DataScope == "By Category" && CurrentRequest.Categories.Any())
            {
                actionMessage.AppendLine("Categories: " + string.Join(", ", CurrentRequest.Categories));
            }

            switch (CurrentRequest.DataAction)
            {
                case "Save File":
                    string fileName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "."
                        + (CurrentRequest.FileFormat ?? "txt").ToLower();
                    string fullFilePath = Path.Combine(ReportDirectoryPath, fileName);

                    if (SaveReportContent(CurrentRequest, fullFilePath))
                    {
                        actionMessage.AppendLine("Success: Report saved as " + CurrentRequest.FileFormat + " to:");
                        actionMessage.AppendLine(ReportDirectoryPath + "\\" + fileName);
                    }
                    else
                    {
                        actionMessage.AppendLine("Failure: Could not save the report file (Check folder access: " + ReportDirectoryPath + ").");
                    }
                    break;
                case "Print":
                    actionMessage.AppendLine("Action: Sending report content to the printer queue.");
                    break;

                case "Display":
                    actionMessage.AppendLine("Action: Preparing report content for on-screen display.");
                    break;
            }

            return actionMessage.ToString();
        }

        private bool SaveReportContent(ReportRequest request, string fullPath)
        {
            if (ReportDirectoryPath == null) return false;
            try
            {
                string content = "Report Generated: " + DateTime.Now.ToString() + "\n" +
                                 "Time Frame: " + request.StartDate.ToShortDateString() + " to " + request.EndDate.ToShortDateString() + "\n" +

                                 "Scope: " + request.DataScope + " (Categories: " + string.Join(", ", request.Categories) + ")\n" +
                                 "Action: Save as " + request.FileFormat;
                File.WriteAllText(fullPath, content);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // *****************************************************************
        // TryParseSingleDate METHOD
        // *****************************************************************
    
         private bool TryParseSingleDate(string input, out DateTime date, out string error)
{
    date = DateTime.MinValue;
    error = null;

    if (string.IsNullOrWhiteSpace(input))
    {
        error = "Date input cannot be empty.";
        return false;
    }

    input = input.Trim();

    // Try exact known formats first
    if (DateTime.TryParseExact(
        input,
        AcceptedDateFormats,
        CultureInfo.InvariantCulture,
        DateTimeStyles.None,
        out date))
    {
        return true;
    }

    // --- FLEXIBLE FALLBACK ---
    // Try a culture-neutral parse (handles “01/10/2025”, “1-10-25”, etc.)
    if (DateTime.TryParse(
        input,
        new CultureInfo("en-GB"), // Use UK order (dd/MM/yyyy)
        DateTimeStyles.None,
        out date))
    {
        return true;
    }

    // Try US fallback just in case
    if (DateTime.TryParse(
        input,
        new CultureInfo("en-US"),
        DateTimeStyles.None,
        out date))
    {
        return true;
    }

    // --- If nothing worked ---
    error = "Invalid input. Please enter a date like 2025-10-01 or 01/10/2025.";
    return false;
}

        private bool TryParseCategories(string input, out List<string> selectedCategories, out string error)
        {
            selectedCategories = new List<string>();
            error = null;
            var rawCategories = input.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(c => c.Trim()).ToList();
            foreach (var category in rawCategories)
            {
                if (ValidReportCategories.Contains(category, StringComparer.OrdinalIgnoreCase))
                {
                    selectedCategories.Add(ValidReportCategories.First(c => c.Equals(category, StringComparison.OrdinalIgnoreCase)));
                }
                else
                {
                    error = "Category '" + category + "' is invalid. Valid: " + string.Join(", ", ValidReportCategories) + ".";
                    return false;
                }
            }
            return selectedCategories.Any();
        }

        public bool FileExistsInReportDir(string input, out string reportName, out string error)
        {
            reportName = null;
            error = null;
            if (ReportDirectoryPath == null)
            {
                error = "Report directory is not accessible."; return false;
            }

            string fullPathToCheck = Path.Combine(ReportDirectoryPath, input);
            if (File.Exists(fullPathToCheck)) { reportName = input; return true; }

            if (File.Exists(fullPathToCheck + ".pdf"))
            {
                reportName = input + ".pdf";
                return true;
            }
            if (File.Exists(fullPathToCheck + ".docx"))
            {
                reportName = input + ".docx";
                return true;
            }
            if (File.Exists(fullPathToCheck + ".txt"))
            {
                reportName = input + ".txt";
                return true;
            }

            error = "Could not find a report file named '" + input + "'.";
            return false;
        }

        public List<string> GetAvailableReports()
        {
            if (ReportDirectoryPath == null || !Directory.Exists(ReportDirectoryPath))
            {
                return new List<string> { "Error: Directory not found."
                };
            }

            try
            {
                var files = Directory.GetFiles(ReportDirectoryPath)
                    .Select(Path.GetFileName)
                    .Where(name => name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||

                               name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                                   name.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    .ToList();
                return files.Any() ? files : new List<string> { "No reports found." };
            }
            catch (Exception ex)
            {
                return new List<string> { "Error listing files: " + ex.Message };
            }
        }
    }
    // ==========================================================
    // 3. NLP Processor
    // ==========================================================
    public class NLPProcessor
    {
        private string _lastHospitalNumber;
        private string _lastIntent;

        private readonly Dictionary<string, string[]> _intentKeywords = new Dictionary<string, string[]>
        {
            { "PatientInfo", new[] { "admit", "patient", "who is", "who's", "details", "info", "profile" } },
            { "MovementInfo", new[] { "move", "transfer", "ward", "shift", "relocate" } },
            { "ReportManager", new[] { "report", "summary", "generate", "create report", "make report" } },
          
            { "OpenReport", new[] { "open report", "show report", "access report", "view report" } },
            { "CountInfo", new[] { "count", "how many", "total", "statistics", "number of" } }
        };
        public string ExtractIntent(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return _lastIntent ??
                "Unknown";
            message = message.ToLower();

            Dictionary<string, int> scores = new Dictionary<string, int>();
            foreach (var kvp in _intentKeywords)
            {
                foreach (var keyword in kvp.Value)
                {
                    if (message.Contains(keyword))
                    {

                        if (scores.ContainsKey(kvp.Key)) scores[kvp.Key] = scores[kvp.Key] + 1;
                        else scores[kvp.Key] = 1;
                    }
                }
            }

            if (scores.Count == 0) return _lastIntent ??
                "Unknown";

            string best = scores.OrderByDescending(s => s.Value).First().Key;
            _lastIntent = best;
            return best;
        }

        public Dictionary<string, string> ExtractEntities(string message)
        {
            var entities = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(message)) return entities;

            message = message.ToLower();

            // Hospital number
            Match hn = Regex.Match(message, @"\b\d{4,8}\b");
            if (hn.Success)
            {
                entities["hospitalNumber"] = hn.Value;
                _lastHospitalNumber = hn.Value;
            }
            else if (!string.IsNullOrEmpty(_lastHospitalNumber))
            {
                entities["hospitalNumber"] = _lastHospitalNumber;
            }

            // Dates (simple YYYY-MM-DD or YYYY/MM/DD)
            MatchCollection dates = Regex.Matches(message, @"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b");
            if (dates.Count == 1) entities["date"] = dates[0].Value;
            else if (dates.Count >= 2)
            {
                entities["startDate"] = dates[0].Value;
                entities["endDate"] = dates[1].Value;
            }

            // Ward
            Match ward = Regex.Match(message, @"ward\s+([a-z0-9\-]+)", RegexOptions.IgnoreCase);
            if (ward.Success) entities["ward"] = ward.Groups[1].Value;

            // Action
            Match action = Regex.Match(message, @"\bsave\b|\bprint\b|\bdisplay\b");
            if (action.Success) entities["action"] = action.Value.Trim();

            // Format
            Match format = Regex.Match(message, @"\bpdf\b|\bdocx\b");
            if (format.Success) entities["format"] = format.Value.ToUpper();

            // Categories
            string[] possibleCategories = new[] { "admission", "deaths", "transfers", "discharges", "consultations" };
            var found = possibleCategories.Where(c => message.Contains(c)).ToList();
            if (found.Any()) entities["categories"] = string.Join(",", found);

            return entities;
        }

        public string SummarizeEntities(Dictionary<string, string> entities)
        {
            if (entities == null || entities.Count == 0) return "(No entities detected)";
            return string.Join("; ", entities.Select(kv => kv.Key + ": " + kv.Value));
        }
    }

    // ==========================================================
    // 4. CHATBOT (Uses NLP + ReportManager + DatabaseHelper + MPHBSMS.CurrentUser)
    // ==========================================================
    public class MphFullFeatureChatBot
    {
        private OleDbConnection _con;
        private NLPProcessor _nlp;
        private ReportManager _reportManager;
        private List<Tuple<string, string>> _conversationHistory;
        public MphFullFeatureChatBot()
        {
            _nlp = new NLPProcessor();
            _reportManager = new ReportManager();
            _conversationHistory = new List<Tuple<string, string>>();

            // Initialize DB (uses your DatabaseHelper implementation)
            try
            {
                DatabaseHelper.InitializeDatabase();
            }
            catch
            {
                // ignore non-fatal init errors here
            }

            _con = new OleDbConnection(DatabaseHelper.ConnectionString);
        }

        private string EnsureConnectionOpen()
        {
            if (_con.State != System.Data.ConnectionState.Open)
            {
                try
                {
                    _con.Open();
                }
                catch (Exception ex)
                {
                    return "Error: could not open database connection. " + ex.Message;
                }
            }
            return null;
        }

        // Main entry: handle a user message and return a textual response
        public string HandleMessage(string userMessage, RichTextBox optionalLogRtb = null)
        {
            // record user in history
            _conversationHistory.Add(Tuple.Create("User", userMessage));
            // use the central MPHBSMS class to get current user
            string currentUser = MPHBSMS.CurrentUser ??
                "UnknownUser";

            // NLP
            string intent = _nlp.ExtractIntent(userMessage);
            var entities = _nlp.ExtractEntities(userMessage);
            string entitySummary = _nlp.SummarizeEntities(entities);

            // Debug log optionally
            if (optionalLogRtb != null)
            {
                ChatLogger.AppendText(optionalLogRtb, "Detected Intent: " + intent, System.Drawing.Color.Blue, true);
                ChatLogger.AppendText(optionalLogRtb, "Entities: " + entitySummary, System.Drawing.Color.DarkGreen, false);
            }

            string response = string.Empty;
            // If the report flow is already started, prioritize that conversation
            if (_reportManager != null && _reportManager.CurrentState != ConversationState.AskForPrimaryAction &&
                _reportManager.CurrentState != ConversationState.Complete)
            {
                // let the report manager state machine handle it

                ProcessResult pr = _reportManager.ProcessAndAdvance(userMessage);
                if (!pr.Success)
                {
                    // FIX: Double newline ensures the next prompt starts on a new line and separates the error
                    response = "⚠ " + pr.ErrorMessage + "\n\n" + _reportManager.GetCurrentPrompt();
                }
                else
                {
                    if (_reportManager.CurrentState == ConversationState.Complete)
                    {
                        response
                            = _reportManager.ExecuteReportLogic();
                    }
                    else
                    {
                        response = _reportManager.GetCurrentPrompt();
                    }
                }

                _conversationHistory.Add(Tuple.Create("Bot", response));
                if (optionalLogRtb != null) ChatLogger.AppendText(optionalLogRtb, "Bot: " + response, System.Drawing.Color.Black, false);
                return response;
            }

            // Otherwise handle intent
            switch (intent)
            {
                case "PatientInfo":
                    response = RespondWithPatientInfo(entities);
                    break;

                case "MovementInfo":
                    response = RespondWithMovementInfo(entities);
                    break;

                case "ReportManager":
                    // Start report flow
                    _reportManager.CurrentState = ConversationState.AskForPrimaryAction;
                    response = "Do you want to generate a new report or open an existing one? (type 'Generate' or 'Open')";
                    break;
                case "OpenReport":
                    // direct open attempt
                    response = HandleOpenReport(userMessage);
                    break;

                case "CountInfo":
                    response = RespondWithCounts(entities);
                    break;

                default:
                    response = "I can help with patients, movements, reports and counts. Example: 'who is patient 12345', 'show movements for 12345', 'generate report'.";
                    break;
            }

            _conversationHistory.Add(Tuple.Create("Bot", response));
            if (optionalLogRtb != null) ChatLogger.AppendText(optionalLogRtb, "Bot: " + response, System.Drawing.Color.Black, false);

            return response;
        }

        private string RespondWithPatientInfo(Dictionary<string, string> entities)
        {
            string hn = entities.ContainsKey("hospitalNumber") ?
                entities["hospitalNumber"] : null;
            if (hn == null) return "Please provide a hospital number (e.g., 'who is patient 12345').";
            string connErr = EnsureConnectionOpen();
            if (connErr != null) return connErr;

            string patientInfo = string.Empty;
            string query = "SELECT name, gender, currentWard, isAdmitted, admissionDate, dischargeDate FROM tblPatientMaster WHERE hospitalNumber = '" + hn + "'";
            try
            {
                using (OleDbCommand cmd = new OleDbCommand(query, _con))
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string name = reader["name"] == DBNull.Value ?
                            "N/A" : reader["name"].ToString();
                        string gender = reader["gender"] == DBNull.Value ? "N/A" : reader["gender"].ToString();
                        string ward = reader["currentWard"] == DBNull.Value ?
                            "N/A" : reader["currentWard"].ToString();
                        string isAdmitted = (reader["isAdmitted"] != DBNull.Value && reader["isAdmitted"].ToString() == "True") ? "Yes" : "No";
                        string admDate = reader["admissionDate"] == DBNull.Value ? "N/A" : Convert.ToDateTime(reader["admissionDate"]).ToString("yyyy-MM-dd");
                        string disDate = reader["dischargeDate"] == DBNull.Value ? "N/A" : Convert.ToDateTime(reader["dischargeDate"]).ToString("yyyy-MM-dd");
                        patientInfo = "Patient: " + name + " (" + gender + "), Current Ward: " + ward + ", Admitted: " + isAdmitted + ", Admission Date: " + admDate + ", Discharge Date: " + disDate;
                    }
                    else
                    {
                        patientInfo = "No patient found for hospital number " + hn + ".";
                    }
                }
            }
            catch (Exception ex)
            {
                patientInfo = "Error querying patient info: " + ex.Message;
            }
            return patientInfo;
        }

        private string RespondWithMovementInfo(Dictionary<string, string> entities)
        {
            string hn = entities.ContainsKey("hospitalNumber") ?
                entities["hospitalNumber"] : null;
            if (hn == null) return "Please provide a hospital number to show movement history.";
            string connErr = EnsureConnectionOpen();
            if (connErr != null) return connErr;
            string query = "SELECT MovementDateTime, fromWard, toWard, category FROM tblPatientMovement WHERE hospitalNumber = '" + hn + "' ORDER BY MovementDateTime DESC";
            List<string> lines = new List<string>();
            try
            {
                using (OleDbCommand cmd = new OleDbCommand(query, _con))
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string dt = reader["MovementDateTime"] == DBNull.Value ?
                            "N/A" : Convert.ToDateTime(reader["MovementDateTime"]).ToString("yyyy-MM-dd HH:mm");
                        string fromWard = reader["fromWard"] == DBNull.Value ? "Unknown" : reader["fromWard"].ToString();
                        string toWard = reader["toWard"] == DBNull.Value ? "Unknown" : reader["toWard"].ToString();
                        string category = reader["category"] == DBNull.Value ? "Transfer" : reader["category"].ToString();
                        lines.Add(dt + " | " + fromWard + " -> " + toWard + " (" + category + ")");
                    }
                }

                if (lines.Count == 0) return "No movement records found for hospital number " + hn + ".";
                return "Movement history for HN " + hn + ":\n" + string.Join("\n", lines);
            }
            catch (Exception ex)
            {
                return "Error querying movements: " + ex.Message;
            }
        }

        private string RespondWithCounts(Dictionary<string, string> entities)
        {
            string connErr = EnsureConnectionOpen();
            if (connErr != null) return connErr;

            try
            {
                int totalPatients = 0;
                int totalMovements = 0;

                using (OleDbCommand cmd = new OleDbCommand("SELECT COUNT(*) FROM tblPatientMaster", _con))
                {
                    object r = cmd.ExecuteScalar();
                    totalPatients = (r != null && r != DBNull.Value) ? Convert.ToInt32(r) : 0;
                }

                using (OleDbCommand cmd = new OleDbCommand("SELECT COUNT(*) FROM tblPatientMovement", _con))
                {
                    object r = cmd.ExecuteScalar();
                    totalMovements = (r != null && r != DBNull.Value) ? Convert.ToInt32(r) : 0;
                }

                return "Total patients: " + totalPatients + ", Total movements: " + totalMovements + ".";
            }
            catch (Exception ex)
            {
                return "Error counting records: " + ex.Message;
            }
        }

        public string HandleOpenReport(string userMessage)
        {
            // try to extract a report name from user message
            string trimmed = userMessage.Replace("open", "").Replace("report", "").Trim();
            if (string.IsNullOrEmpty(trimmed)) trimmed = userMessage.Trim();

            string foundName;
            string error;
            if (_reportManager.FileExistsInReportDir(trimmed, out foundName, out error))
            {
                _reportManager.CurrentRequest.ReportName = foundName;
                _reportManager.CurrentRequest.DataAction = "OpenFile";
                _reportManager.CurrentState = ConversationState.ConfirmationAndGenerate;
                return "Ready to open report " + foundName + ". Type 'Confirm' to proceed.";
            }

            List<string> avail = _reportManager.GetAvailableReports();
            string availText = string.Join("\n - ", avail);
            return error + "\nAvailable reports:\n - " + availText;
        }

        // Expose conversation history for UI usage
        public List<Tuple<string, string>> GetConversationHistory()
        {
            return _conversationHistory;
        }
    }
}