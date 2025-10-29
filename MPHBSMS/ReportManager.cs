using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// ==========================================================
// 1. HELPER DEFINITIONS
// ==========================================================

public enum ConversationState
{
    Greeting,
    AskForPrimaryAction,
    AskForTimeFrame,
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

    private static readonly string[] AcceptedDateFormats = new string[]
    {
        "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy", "M/d/yyyy", "dd-MMM-yyyy"
    };
    private readonly List<string> ValidReportCategories = new List<string>
    {
        "Admission", "Deaths", "Transfers", "Discharges", "Consultations"
    };

    // --- Constructor ---
    public ReportManager()
    {
        CurrentRequest = new ReportRequest();
        _currentState = ConversationState.Greeting;
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
        bool isValid = ProcessInput(input, out errorMessage);

        if (isValid)
        {
            MoveToNextState();
        }

        return new ProcessResult { Success = isValid, ErrorMessage = errorMessage };
    }

    private bool ProcessInput(string input, out string errorMessage)
    {
        input = input.Trim();
        errorMessage = null;
        bool isValid = false;

        DateTime startDate;
        DateTime endDate;
        List<string> selectedCategories;
        string reportName;


        switch (CurrentState)
        {
            case ConversationState.AskForPrimaryAction:
                // Added friendly synonyms: "Generate", "Open"
                if (input.Equals("GenerateReport", StringComparison.OrdinalIgnoreCase) || input.Equals("Generate", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.PrimaryAction = "GenerateReport"; isValid = true; }
                else if (input.Equals("OpenExisting", StringComparison.OrdinalIgnoreCase) || input.Equals("Open", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.PrimaryAction = "OpenExisting"; isValid = true; }
                else { errorMessage = "Please enter 'GenerateReport' or 'OpenExisting'."; }
                break;

            case ConversationState.AskForTimeFrame:
                if (TryParseTimeFrame(input, out startDate, out endDate, out errorMessage))
                { CurrentRequest.StartDate = startDate; CurrentRequest.EndDate = endDate; isValid = true; }
                break;

            case ConversationState.AskForDataScope:
                // Added friendly synonyms: "All"
                if (input.Equals("All Data", StringComparison.OrdinalIgnoreCase) || input.Equals("All", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.DataScope = "All Data"; CurrentRequest.Categories.Clear(); isValid = true; }
                else
                {
                    if (TryParseCategories(input, out selectedCategories, out errorMessage))
                    { CurrentRequest.DataScope = "By Category"; CurrentRequest.Categories = selectedCategories; isValid = true; }
                }
                break;

            case ConversationState.AskForDataAction:
                // Added friendly synonyms: "Save", "Print", "Display"
                if (input.Equals("Save File", StringComparison.OrdinalIgnoreCase) || input.Equals("Save", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.DataAction = "Save File"; isValid = true; }
                else if (input.Equals("Print", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.DataAction = "Print"; isValid = true; }
                else if (input.Equals("Display", StringComparison.OrdinalIgnoreCase))
                { CurrentRequest.DataAction = "Display"; isValid = true; }
                else { errorMessage = "Please enter 'Save File', 'Print', or 'Display'."; }
                break;

            case ConversationState.AskForSaveFormat:
                if (input.Equals("PDF", StringComparison.OrdinalIgnoreCase)) { CurrentRequest.FileFormat = "PDF"; isValid = true; }
                else if (input.Equals("DOCX", StringComparison.OrdinalIgnoreCase)) { CurrentRequest.FileFormat = "DOCX"; isValid = true; }
                else { errorMessage = "Please enter 'PDF' or 'DOCX'."; }
                break;

            case ConversationState.AskForReportToOpen:
                if (FileExistsInReportDir(input, out reportName, out errorMessage))
                { CurrentRequest.ReportName = reportName; isValid = true; }
                break;

            case ConversationState.ConfirmationAndGenerate:
                // Added friendly synonyms: "Yes", "Go"
                if (input.Equals("Confirm", StringComparison.OrdinalIgnoreCase) || input.Equals("Yes", StringComparison.OrdinalIgnoreCase) || input.Equals("Go", StringComparison.OrdinalIgnoreCase))
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
            case ConversationState.AskForPrimaryAction:
                CurrentState = (CurrentRequest.PrimaryAction == "GenerateReport") ? ConversationState.AskForTimeFrame : ConversationState.AskForReportToOpen;
                break;
            case ConversationState.AskForTimeFrame: CurrentState = ConversationState.AskForDataScope; break;
            case ConversationState.AskForDataScope: CurrentState = ConversationState.AskForDataAction; break;
            case ConversationState.AskForDataAction:
                CurrentState = (CurrentRequest.DataAction == "Save File") ? ConversationState.AskForSaveFormat : ConversationState.ConfirmationAndGenerate;
                break;
            case ConversationState.AskForSaveFormat: CurrentState = ConversationState.ConfirmationAndGenerate; break;
            case ConversationState.AskForReportToOpen: CurrentState = ConversationState.ConfirmationAndGenerate; break;
            case ConversationState.ConfirmationAndGenerate: CurrentState = ConversationState.Complete; break;
        }
    }

    public string GetCurrentPrompt()
    {
        switch (CurrentState)
        {
            case ConversationState.Greeting: return "Welcome! Ready for your first command.";
            case ConversationState.AskForPrimaryAction: return "Do you want to **GenerateReport** or **OpenExisting**? (Tip: Try 'Generate' or 'Open')";
            case ConversationState.AskForTimeFrame: return "What is the time frame? (e.g., '2023-01-01 to 2023-10-31')";
            case ConversationState.AskForDataScope:
                return "Do you want **All Data** or specific categories? (Valid: " + string.Join(", ", ValidReportCategories) + " or 'All')";
            case ConversationState.AskForDataAction: return "What do you want to do with the data? (**Save File**, **Print**, or **Display**)";
            case ConversationState.AskForSaveFormat: return "What format do you want to save as? (**PDF** or **DOCX**)";
            case ConversationState.AskForReportToOpen:
                string list = string.Join("\n - ", GetAvailableReports());
                return "What is the name of the existing report you want to open?\n\nAvailable Reports:\n - " + list;
            case ConversationState.ConfirmationAndGenerate: return "Ready to proceed with " + CurrentRequest.PrimaryAction + ". Please confirm.";
            case ConversationState.Complete: return "Operation complete. Thank you!";
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
        string actionMessage = "Report Generation for scope: **" + CurrentRequest.DataScope + "** (Dates: " + CurrentRequest.StartDate.ToShortDateString() + " to " + CurrentRequest.EndDate.ToShortDateString() + ")\n";

        switch (CurrentRequest.DataAction)
        {
            case "Save File":
                string fileName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + CurrentRequest.FileFormat.ToLower();
                string fullFilePath = Path.Combine(ReportDirectoryPath, fileName);

                if (SaveReportContent(CurrentRequest, fullFilePath))
                {
                    actionMessage += "Success: Report saved as **" + CurrentRequest.FileFormat + "** to:\n" + ReportDirectoryPath + "\\" + fileName;
                }
                else
                {
                    actionMessage += "Failure: Could not save the report file.";
                }
                break;

            case "Print":
                actionMessage += "Action: Sending report content to the printer queue.";
                break;

            case "Display":
                actionMessage += "Action: Preparing report content for on-screen display.";
                break;
        }

        return actionMessage;
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

    private bool TryParseTimeFrame(string input, out DateTime startDate, out DateTime endDate, out string error)
    {
        startDate = DateTime.MinValue; endDate = DateTime.MinValue; error = null;

        string cleanedInput = input.Replace("to", "|").Trim();
        string[] parts = cleanedInput.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length != 2)
        {
            error = "Please enter two distinct dates separated by 'to'.";
            return false;
        }

        string startDateString = parts[0].Trim();
        string endDateString = parts[1].Trim();

        if (!DateTime.TryParseExact(startDateString, AcceptedDateFormats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out startDate))
        {
            error = "Invalid start date format. Ensure it's a valid date (e.g., YYYY-MM-DD or MM/DD/YYYY).";
            return false;
        }

        if (!DateTime.TryParseExact(endDateString, AcceptedDateFormats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out endDate))
        {
            error = "Invalid end date format. Ensure it's a valid date.";
            return false;
        }

        if (startDate > endDate)
        {
            error = "The start date cannot be after the end date.";
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
            { selectedCategories.Add(category); }
            else
            { error = "Category '" + category + "' is invalid. Valid: " + string.Join(", ", ValidReportCategories) + "."; return false; }
        }
        return true;
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