using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;

namespace MPHBSMS
{
    // ==========================================================
    // 1. HELPER DEFINITIONS
    // ==========================================================

    public enum ConversationState
    {
        Greeting,
        AskForPrimaryAction,
        AskForStartDate, // Separated state
        AskForEndDate,   // Separated state
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
        public System.DateTime StartDate { get; set; }
        public System.DateTime EndDate { get; set; }
        public string DataScope { get; set; }
        public string DataAction { get; set; }
        public string FileFormat { get; set; }
        public string ReportName { get; set; }

        public List<string> Categories
        {
            get { return _categories; }
            set { _categories = value; }
        }

        public ReportRequest()
        {
            _categories = new List<string>();
        }
    }

    public class ProcessResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }


    // ==========================================================
    // 2. REPORT MANAGER CLASS
    // ==========================================================

    public class ReportManager
    {
        // --- Fields and Properties ---
        private ConversationState _currentState;
        public ConversationState CurrentState
        {
            get { return _currentState; }
            set { _currentState = value; }
        }

        public ReportRequest CurrentRequest { get; private set; }

        private const string AppFolderName = "MyReportGeneratorFiles";
        private readonly string ReportDirectoryPath;

        // 🌟 FINAL FIX: Comprehensive list of date formats to accept all user inputs.
        private static readonly string[] AcceptedDateFormats = new string[]
        {
            "yyyy-MM-dd", "yyyy/MM/dd", 
            "MM-dd-yyyy", "MM/dd/yyyy", 
            "M-d-yyyy", "M/d/yyyy",     
            "dd-MM-yyyy", "dd/MM/yyyy", 
            "d-M-yyyy", "d/M/yyyy",     
            "dd-MMM-yyyy"
        };

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

        // --- Core State Machine ---

        public ProcessResult ProcessAndAdvance(string input)
        {
            string errorMessage = null;
            // CRITICAL SANITIZATION: Clean up any hidden control characters and trim
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
                    MoveToNextState();
                }
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


            switch (CurrentState)
            {
                case ConversationState.Greeting:
                case ConversationState.AskForPrimaryAction:
                    string lowerInput = cleanedInput.ToLower();

                    if (lowerInput.Equals("generatereport") || lowerInput.Equals("generate") || lowerInput.Contains("generate"))
                    { CurrentRequest.PrimaryAction = "GenerateReport"; isValid = true; }
                    else if (lowerInput.Equals("openexisting") || lowerInput.Equals("open") || lowerInput.Contains("open"))
                    { CurrentRequest.PrimaryAction = "OpenExisting"; isValid = true; }
                    else { errorMessage = "Please enter 'GenerateReport' or 'OpenExisting'."; }
                    break;

                case ConversationState.AskForStartDate:
                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    { CurrentRequest.StartDate = parsedDate; isValid = true; }
                    break;

                case ConversationState.AskForEndDate:
                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    {
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
                    break;

                case ConversationState.AskForDataScope:
                    if (cleanedInput.Equals("All Data", StringComparison.OrdinalIgnoreCase) || cleanedInput.Equals("All", StringComparison.OrdinalIgnoreCase))
                    { CurrentRequest.DataScope = "All Data"; CurrentRequest.Categories.Clear(); isValid = true; }
                    else
                    {
                        if (TryParseCategories(cleanedInput, out selectedCategories, out errorMessage))
                        { CurrentRequest.DataScope = "By Category"; CurrentRequest.Categories = selectedCategories; isValid = true; }
                    }
                    break;

                case ConversationState.AskForDataAction:
                    if (cleanedInput.Equals("Save File", StringComparison.OrdinalIgnoreCase) || cleanedInput.Equals("Save", StringComparison.OrdinalIgnoreCase))
                    { CurrentRequest.DataAction = "Save File"; isValid = true; }
                    else if (cleanedInput.Equals("Print", StringComparison.OrdinalIgnoreCase))
                    { CurrentRequest.DataAction = "Print"; isValid = true; }
                    else if (cleanedInput.Equals("Display", StringComparison.OrdinalIgnoreCase))
                    { CurrentRequest.DataAction = "Display"; isValid = true; }
                    else { errorMessage = "Please enter 'Save File', 'Print', or 'Display'."; }
                    break;

                case ConversationState.AskForSaveFormat:
                    if (cleanedInput.Equals("PDF", StringComparison.OrdinalIgnoreCase)) { CurrentRequest.FileFormat = "PDF"; isValid = true; }
                    else if (cleanedInput.Equals("DOCX", StringComparison.OrdinalIgnoreCase)) { CurrentRequest.FileFormat = "DOCX"; isValid = true; }
                    else { errorMessage = "Please enter 'PDF' or 'DOCX'."; }
                    break;

                case ConversationState.AskForReportToOpen:
                    if (FileExistsInReportDir(cleanedInput, out reportName, out errorMessage))
                    { CurrentRequest.ReportName = reportName; isValid = true; }
                    break;

                case ConversationState.ConfirmationAndGenerate:
                    if (cleanedInput.Equals("Confirm", StringComparison.OrdinalIgnoreCase) || cleanedInput.Equals("Yes", StringComparison.OrdinalIgnoreCase) || cleanedInput.Equals("Go", StringComparison.OrdinalIgnoreCase))
                    { isValid = true; }
                    else { errorMessage = "Please type 'Confirm' to proceed with generation."; }
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
                    CurrentState = (CurrentRequest.PrimaryAction == "GenerateReport") ? ConversationState.AskForStartDate : ConversationState.AskForReportToOpen;
                    break;
                case ConversationState.AskForStartDate: CurrentState = ConversationState.AskForEndDate; break;
                case ConversationState.AskForEndDate: CurrentState = ConversationState.AskForDataScope; break;
                case ConversationState.AskForDataScope: CurrentState = ConversationState.AskForDataAction; break;
                case ConversationState.AskForDataAction:
                    CurrentState = (CurrentRequest.DataAction == "Save File") ? ConversationState.AskForSaveFormat : ConversationState.ConfirmationAndGenerate;
                    break;
                case ConversationState.AskForSaveFormat: CurrentState = ConversationState.ConfirmationAndGenerate; break;
                case ConversationState.AskForReportToOpen: CurrentState = ConversationState.ConfirmationAndGenerate; break;
            }
        }

        public string GetCurrentPrompt()
        {
            switch (CurrentState)
            {
                case ConversationState.Greeting:
                    return "Welcome to the Report Generator! I'm here to guide you through creating or opening a report.";
                case ConversationState.AskForPrimaryAction:
                    return "Do you want to **GenerateReport** or **OpenExisting**? (Tip: Try 'Generate' or 'Open')";
                case ConversationState.AskForStartDate: return "What is the **Start Date** for the report? (e.g., '2023-01-01' or '01/01/2023')";
                case ConversationState.AskForEndDate: return "What is the **End Date** for the report? (e.g., '2023-10-31' or '10/31/2023')";

                case ConversationState.AskForDataScope:
                    return "Do you want **All Data** or specific categories? (Valid: " + string.Join(", ", ValidReportCategories) + " or 'All')";
                case ConversationState.AskForDataAction: return "What do you want to do with the data? (**Save File**, **Print**, or **Display**)";
                case ConversationState.AskForSaveFormat: return "What format do you want to save as? (**PDF** or **DOCX**)";
                case ConversationState.AskForReportToOpen:
                    string list = string.Join("\n - ", GetAvailableReports());
                    return "What is the name of the existing report you want to open?\n\nAvailable Reports:\n - " + list;
                case ConversationState.ConfirmationAndGenerate: return "Ready to proceed with " + CurrentRequest.PrimaryAction + ". Please confirm. (Type 'Confirm')";
                case ConversationState.Complete: return "Operation complete. Type **Generate** or **Open** to start a new report.";
                default: return "Error: Unknown state.";
            }
        }

        // --- Execution and Helpers ---

        public string ExecuteReportLogic()
        {
            if (CurrentRequest.PrimaryAction == "OpenExisting")
            {
                return "Action Complete: Attempting to open existing report file: **" + CurrentRequest.ReportName + "**";
            }

            // Logic for GENERATE REPORT
            StringBuilder actionMessage = new StringBuilder();
            actionMessage.AppendLine("Report Generation for scope: **" + CurrentRequest.DataScope + "** (Dates: " + CurrentRequest.StartDate.ToShortDateString() + " to " + CurrentRequest.EndDate.ToShortDateString() + ")");

            if (CurrentRequest.DataScope == "By Category" && CurrentRequest.Categories.Any())
            {
                actionMessage.AppendLine("Categories: " + string.Join(", ", CurrentRequest.Categories));
            }

            switch (CurrentRequest.DataAction)
            {
                case "Save File":
                    string fileName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + CurrentRequest.FileFormat.ToLower();
                    string fullFilePath = Path.Combine(ReportDirectoryPath, fileName);

                    if (SaveReportContent(CurrentRequest, fullFilePath))
                    {
                        actionMessage.AppendLine("Success: Report saved as **" + CurrentRequest.FileFormat + "** to:");
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

        private bool TryParseSingleDate(string input, out DateTime date, out string error)
        {
            date = DateTime.MinValue; error = null;

            if (!DateTime.TryParseExact(
                input,
                AcceptedDateFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date))
            {
                error = "Invalid date format. Please use a valid format (e.g., YYYY-MM-DD or MM/DD/YYYY).";
                return false;
            }

            return true;
        }

        private bool TryParseCategories(string input, out List<string> selectedCategories, out string error)
        {
            selectedCategories = new List<string>(); error = null;
            var rawCategories = input.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(c => c.Trim())
                                     .ToList();

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

        private bool FileExistsInReportDir(string input, out string reportName, out string error)
        {
            reportName = null; error = null;
            if (ReportDirectoryPath == null) { error = "Report directory is not accessible."; return false; }

            string fullPathToCheck = Path.Combine(ReportDirectoryPath, input);
            if (File.Exists(fullPathToCheck)) { reportName = input; return true; }

            if (File.Exists(fullPathToCheck + ".pdf")) { reportName = input + ".pdf"; return true; }
            if (File.Exists(fullPathToCheck + ".docx")) { reportName = input + ".docx"; return true; }

            error = "Could not find a report file named '" + input + "'.";
            return false;
        }

        private List<string> GetAvailableReports()
        {
            if (ReportDirectoryPath == null || !Directory.Exists(ReportDirectoryPath))
            {
                return new List<string> { "Error: Directory not found." };
            }

            try
            {
                var files = Directory.GetFiles(ReportDirectoryPath)
                    .Select(Path.GetFileName)
                    .Where(name => name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                                   name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                return files.Any() ? files : new List<string> { "No reports found." };
            }
            catch (Exception ex)
            {
                return new List<string> { "Error listing files: " + ex.Message };
            }
        }
    }
}