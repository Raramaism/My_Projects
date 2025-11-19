using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace MPHBSMS
{
    public enum ConversationState
    {
        Greeting,
        AskForPrimaryAction,
        AskForStartDate,
        AskForStartMonth,
        AskForStartYear,
        AskForEndDate,
        AskForEndMonth,
        AskForEndYear,
        AskForDataScope,
        AskForDataAction,
        AskForSaveFormat,
        AskForReportToOpen,
        ConfirmationAndGenerate,
        Complete
    }

    public class ReportRequest
    {
        public string PrimaryAction { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string DataScope { get; set; }
        public string DataAction { get; set; }
        public string FileFormat { get; set; }
        public string ReportName { get; set; }

        public int? TempDay { get; set; }
        public int? TempMonth { get; set; }
        public int? TempYear { get; set; }

        public List<string> Categories { get; set; }

        public ReportRequest()
        {
            Categories = new List<string>();
            StartDate = DateTime.Now.Date;
            EndDate = DateTime.Now.Date;
        }

        public void ClearTempDate()
        {
            TempDay = null;
            TempMonth = null;
            TempYear = null;
        }

        public string StartDateString
        {
            get { return StartDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture); }
        }

        public string EndDateString
        {
            get { return EndDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture); }
        }
    }

    public class ProcessResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class ReportManager
    {
        private ConversationState _currentState;
        public ConversationState CurrentState
        {
            get { return _currentState; }
            set { _currentState = value; }
        }

        public ReportRequest CurrentRequest { get; private set; }

        private const string AppFolderName = "MyReportGeneratorFiles";
        private readonly string ReportDirectoryPath;

        private readonly List<string> ValidReportCategories = new List<string>
        {
            "Admission", "Deaths", "Transfers", "Discharges", "Consultations"
        };

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

        public ProcessResult ProcessAndAdvance(string input)
        {
            string errorMessage = null;
            if (input == null) input = string.Empty;

            string cleanedInput = new string(input.Where(c => !char.IsControl(c)).ToArray()).Trim();
            bool isValid = ProcessInput(cleanedInput, out errorMessage);

            if (isValid)
            {
                if (CurrentState == ConversationState.ConfirmationAndGenerate)
                    CurrentState = ConversationState.Complete;
                else
                    MoveToNextState();
            }
            else if (errorMessage == "NEXT_STEP_OK")
            {
                return new ProcessResult { Success = true, ErrorMessage = null };
            }

            return new ProcessResult { Success = isValid, ErrorMessage = errorMessage };
        }

        private bool ProcessInput(string cleanedInput, out string errorMessage)
        {
            errorMessage = null;
            bool isValid = false;
            DateTime parsedDate;
            int parsedNumber;

            switch (CurrentState)
            {
                case ConversationState.Greeting:
                case ConversationState.AskForPrimaryAction:
                    string lowerInput = cleanedInput.ToLower();
                    if (lowerInput.Contains("generate") || lowerInput.Contains("create"))
                    {
                        CurrentRequest.PrimaryAction = "GenerateReport";
                        isValid = true;
                    }
                    else if (lowerInput.Contains("open"))
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
                    CurrentRequest.ClearTempDate();
                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    {
                        CurrentRequest.StartDate = parsedDate;
                        isValid = true;
                    }
                    else if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 31)
                    {
                        CurrentRequest.TempDay = parsedNumber;
                        CurrentState = ConversationState.AskForStartMonth;
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(errorMessage))
                            errorMessage = "Invalid input. Please enter a date like 2025-10-01 or 01/10/2025.";
                    }
                    break;

                case ConversationState.AskForStartMonth:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 12)
                    {
                        CurrentRequest.TempMonth = parsedNumber;
                        CurrentState = ConversationState.AskForStartYear;
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        errorMessage = "Invalid month. Enter number 1 to 12.";
                    }
                    break;

                case ConversationState.AskForStartYear:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1900 && parsedNumber <= DateTime.Now.Year + 5)
                    {
                        int year = parsedNumber;
                        int month = CurrentRequest.TempMonth ?? 1;
                        int day = CurrentRequest.TempDay ?? 1;

                        try
                        {
                            CurrentRequest.StartDate = new DateTime(year, month, day);
                            CurrentRequest.ClearTempDate();
                            isValid = true;
                        }
                        catch
                        {
                            errorMessage = "Invalid date combination (" + day + "/" + month + "/" + year + ").";
                        }
                    }
                    else
                    {
                        errorMessage = "Invalid year. Enter a 4-digit year.";
                    }
                    break;

                case ConversationState.AskForEndDate:
                    CurrentRequest.ClearTempDate();
                    if (TryParseSingleDate(cleanedInput, out parsedDate, out errorMessage))
                    {
                        if (parsedDate < CurrentRequest.StartDate)
                        {
                            errorMessage = "End date cannot be before start date (" + CurrentRequest.StartDateString + ").";
                        }
                        else
                        {
                            CurrentRequest.EndDate = parsedDate;
                            isValid = true;
                        }
                    }
                    else if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 31)
                    {
                        CurrentRequest.TempDay = parsedNumber;
                        CurrentState = ConversationState.AskForEndMonth;
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(errorMessage))
                            errorMessage = "Invalid input. Please enter a date like 2025-10-01 or 01/10/2025.";
                    }
                    break;

                case ConversationState.AskForEndMonth:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 12)
                    {
                        CurrentRequest.TempMonth = parsedNumber;
                        CurrentState = ConversationState.AskForEndYear;
                        errorMessage = "NEXT_STEP_OK";
                    }
                    else
                    {
                        errorMessage = "Invalid month. Enter number 1 to 12.";
                    }
                    break;

                case ConversationState.AskForEndYear:
                    if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1900 && parsedNumber <= DateTime.Now.Year + 5)
                    {
                        int year = parsedNumber;
                        int month = CurrentRequest.TempMonth ?? 1;
                        int day = CurrentRequest.TempDay ?? 1;

                        try
                        {
                            parsedDate = new DateTime(year, month, day);
                            if (parsedDate < CurrentRequest.StartDate)
                                errorMessage = "End date cannot be before start date (" + CurrentRequest.StartDateString + ").";
                            else
                            {
                                CurrentRequest.EndDate = parsedDate;
                                CurrentRequest.ClearTempDate();
                                isValid = true;
                            }
                        }
                        catch
                        {
                            errorMessage = "Invalid end date combination.";
                        }
                    }
                    else
                    {
                        errorMessage = "Invalid year. Enter a 4-digit year.";
                    }
                    break;
            }

            return isValid;
        }

        private bool TryParseSingleDate(string input, out DateTime date, out string error)
        {
            date = DateTime.MinValue;
            error = null;

            if (string.IsNullOrWhiteSpace(input))
            {
                error = "Date input cannot be empty.";
                return false;
            }

            // Keep original and also a normalized version (all separators -> '-')
            string original = input.Trim();
            string normalized = original.Replace('\\', '-')
                                        .Replace('/', '-')
                                        .Replace('.', '-')
                                        .Replace('_', '-')
                                        .Trim();

            // Common patterns (with both '-' and '/' versions included)
            string[] formats = new string[]
    {
        "yyyy-MM-dd", "yyyy/MM/dd",
        "dd-MM-yyyy", "dd/MM/yyyy",
        "MM-dd-yyyy", "MM/dd/yyyy",
        "d-M-yyyy",   "d/M/yyyy",
        "M-d-yyyy",   "M/d/yyyy",
        "yyyy-M-d",
        "dd-MMM-yyyy",
        "yyyyMMdd",
        "ddMMyyyy"
    };

            // 1) Try exact parse on the normalized string with invariant culture
            foreach (string fmt in formats)
            {
                // try format as-is on normalized (which uses '-')
                string fmtNormalized = fmt.Replace('/', '-');
                if (DateTime.TryParseExact(normalized, fmtNormalized, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    return true;

                // try format as-is on original (preserves '/' if present)
                if (DateTime.TryParseExact(original, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    return true;
            }

            // 2) Try broad parsing with several cultures (lenient)
            if (DateTime.TryParse(normalized, new CultureInfo("en-GB"), DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(normalized, new CultureInfo("en-US"), DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(original, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(original, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                return true;

            // final fallback: try removing extra spaces and try parse again
            string compact = new string(original.Where(c => !char.IsWhiteSpace(c)).ToArray());
            if (DateTime.TryParse(compact, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                return true;

            error = "Invalid date. Please enter a valid date like 2025-10-01 or 01/10/2025.";
            return false;
        }

        private void MoveToNextState()
        {
            switch (CurrentState)
            {
                case ConversationState.AskForPrimaryAction:
                    CurrentState = (CurrentRequest.PrimaryAction == "GenerateReport")
                        ? ConversationState.AskForStartDate
                        : ConversationState.AskForReportToOpen;
                    break;
                case ConversationState.AskForStartDate:
                case ConversationState.AskForStartYear:
                    CurrentState = ConversationState.AskForEndDate;
                    break;
                case ConversationState.AskForEndDate:
                case ConversationState.AskForEndYear:
                    CurrentState = ConversationState.AskForDataScope;
                    break;
            }
        }

        // === IMPLEMENTED REPORT GENERATION ===
        public string ExecuteReportLogic()
        {
            if (CurrentRequest.PrimaryAction == "OpenExisting")
            {
                return "Action Complete: Attempting to open existing report file: " + CurrentRequest.ReportName;
            }

            StringBuilder actionMessage = new StringBuilder();
            actionMessage.AppendLine("Report Generation for scope: " + CurrentRequest.DataScope +
                                     " (Dates: " + CurrentRequest.StartDateString +
                                     " to " + CurrentRequest.EndDateString + ")");

            if (CurrentRequest.DataScope == "By Category" && CurrentRequest.Categories.Any())
            {
                actionMessage.AppendLine("Categories: " + string.Join(", ", CurrentRequest.Categories));
            }

            // Choose file format
            string format = (CurrentRequest.FileFormat ?? "txt").ToLower();
            string fileName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + format;
            string fullFilePath = Path.Combine(ReportDirectoryPath, fileName);

            if (SaveReportContent(CurrentRequest, fullFilePath))
            {
                actionMessage.AppendLine("Success: Report saved as " + format.ToUpper() + " to:");
                actionMessage.AppendLine(fullFilePath);
            }
            else
            {
                actionMessage.AppendLine("Failure: Could not save the report file (Check folder access: " + ReportDirectoryPath + ").");
            }

            return actionMessage.ToString();
        }

        private bool SaveReportContent(ReportRequest request, string fullPath)
        {
            if (ReportDirectoryPath == null) return false;
            try
            {
                string content = "REPORT GENERATED ON: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\n" +
                                 "PERIOD: " + request.StartDateString + " to " + request.EndDateString + "\n" +
                                 "DATA SCOPE: " + request.DataScope + "\n" +
                                 "FORMAT: " + request.FileFormat + "\n" +
                                 "CATEGORIES: " + string.Join(", ", request.Categories);

                File.WriteAllText(fullPath, content, Encoding.UTF8);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string GetCurrentPrompt()
        {
            switch (CurrentState)
            {
                case ConversationState.AskForPrimaryAction:
                    return "Do you want to Generate or Open an existing report?";
                case ConversationState.AskForStartDate:
                    return "What is the Start Date for the report? (e.g., 2025-10-01 or 01/10/2025)";
                case ConversationState.AskForStartMonth:
                    return "Please enter the Month number for the start date (1-12).";
                case ConversationState.AskForStartYear:
                    return "Please enter the Year for the start date (e.g., 2024).";
                case ConversationState.AskForEndDate:
                    return "What is the End Date for the report? (e.g., 2025-12-31 or 31/12/2025)";
                case ConversationState.AskForEndMonth:
                    return "Please enter the Month number for the end date (1-12).";
                case ConversationState.AskForEndYear:
                    return "Please enter the Year for the end date (e.g., 2024).";
                default:
                    return "Continue with the next step.";
            }
        }
    }
}
