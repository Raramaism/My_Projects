using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace MPHBSMS
{
    // Single self-contained class as requested
    /// <summary>
    /// Manages the state and logic for conversational report generation and database interaction.
    /// It includes connection logic for an Access database (using OleDb) and file handling.
    /// </summary>
    public class ReportManager
    {
        // ----- Conversation state machine (nested enum) -----
        /// <summary>
        /// Defines the steps in the conversational workflow for report generation.
        /// </summary>
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

        // ----- Nested DTOs -----
        /// <summary>
        /// Data Transfer Object to hold the current request parameters.
        /// </summary>
        public class ReportRequest
        {
            public string PrimaryAction { get; set; }
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            public string DataScope { get; set; } // normalized category or "ALL"
            public string DataAction { get; set; } // ward or "user:username" or other filter
            public string FileFormat { get; set; } // txt/docx/xlsx
            public string ReportName { get; set; }

            // Temporary fields for multi-step date input
            public int? TempDay { get; set; }
            public int? TempMonth { get; set; }
            public int? TempYear { get; set; }

            public List<string> Categories { get; set; }

            public ReportRequest()
            {
                Categories = new List<string>();
                StartDate = DateTime.Now.Date;
                EndDate = DateTime.Now.Date;
                FileFormat = "txt";
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

        /// <summary>
        /// Result of processing a user input.
        /// </summary>
        public class ProcessResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        // ----- Fields & constants -----
        private ConversationState _currentState;
        public ConversationState CurrentState
        {
            get { return _currentState; }
            set { _currentState = value; }
        }

        public ReportRequest CurrentRequest { get; private set; }

        private const string AppFolderName = "MyReportGeneratorFiles";
        private readonly string ReportDirectoryPath;

        // Database configuration (uses DataDirectory set during initialization)
        public static string ConnectionString
        {
            get
            {
                // Note: |DataDirectory| is dynamically set to the AppData folder for persistence
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MPHDATABASE.accdb;Persist Security Info=False;";
            }
        }

        // Canonical categories used in database (tblPatientMovement)
        private const string CatAdmission = "Admission";
        private const string CatTransferIn = "InterWardTransferIn";
        private const string CatTransferOut = "InterWardTransferOut";
        private const string CatDischarge = "Discharge";
        private const string CatDeath = "Death";
        private const string CatConsultations = "Consultations";

        // Ward list provided for filtering/validation
        private readonly List<string> WardList = new List<string>
        {
            "Mental Health Unit","Female Ward","Paedatric Ward","Male Ward","PostNatal Ward",
            "NeoNatal Ward","AnteNatal Ward","Labor Ward","Accident and Emergence","Mortuary"
        };

        // Valid categories (for UI or validation)
        private readonly List<string> ValidReportCategories;

        // ----- Constructor -----
        public ReportManager()
        {
            CurrentRequest = new ReportRequest();
            _currentState = ConversationState.AskForPrimaryAction;
            ReportDirectoryPath = InitializeReportDirectory();

            ValidReportCategories = new List<string>
            {
                CatAdmission, CatDeath, CatTransferIn, CatTransferOut, CatDischarge, CatConsultations
            };
        }

        // ----- Database initialization (copies accdb to AppData and sets DataDirectory) -----
        /// <summary>
        /// Initializes the database path by copying the main Access DB file
        /// to the application's local AppData folder and setting the |DataDirectory| macro.
        /// </summary>
        public static void InitializeDatabase()
        {
            try
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "MaronderaProvincialHospitalSystem");

                if (!Directory.Exists(appDataPath))
                    Directory.CreateDirectory(appDataPath);

                string sourceDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MPHDATABASE.accdb");
                string targetDb = Path.Combine(appDataPath, "MPHDATABASE.accdb");

                if (!File.Exists(targetDb))
                {
                    // Copy only if source exists
                    if (File.Exists(sourceDb))
                        File.Copy(sourceDb, targetDb);
                }

                AppDomain.CurrentDomain.SetData("DataDirectory", appDataPath);
            }
            catch
            {
                // Swallow exceptions during initialization; connection attempt will fail later if needed.
            }
        }

        // ----- UI file folder initialization -----
        /// <summary>
        /// Creates the dedicated folder in My Documents for saving reports.
        /// </summary>
        /// <returns>The full path to the report directory.</returns>
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

        // ----- Conversation processing -----
        /// <summary>
        /// Processes the user's input based on the current state and advances the conversation.
        /// </summary>
        /// <param name="input">The user's text input.</param>
        /// <returns>A ProcessResult indicating success and any error message.</returns>
        public ProcessResult ProcessAndAdvance(string input)
        {
            string errorMessage = null;
            if (input == null) input = string.Empty;

            // Remove control characters (e.g., from copy-paste)
            string cleanedInput = new string(input.Where(c => !char.IsControl(c)).ToArray()).Trim();
            bool isValid = ProcessInput(cleanedInput, out errorMessage);

            if (isValid)
            {
                // Only move state if input was valid and not a temporary date step
                if (CurrentState == ConversationState.ConfirmationAndGenerate)
                    CurrentState = ConversationState.Complete;
                else
                    MoveToNextState();
            }
            else if (errorMessage == "NEXT_STEP_OK")
            {
                // This is a special flag used when switching to TempDay/TempMonth/TempYear
                return new ProcessResult { Success = true, ErrorMessage = null };
            }

            return new ProcessResult { Success = isValid, ErrorMessage = errorMessage };
        }

        /// <summary>
        /// Handles the validation and storage of the input for the current state.
        /// </summary>
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
                    string lowerInput = cleanedInput.ToLowerInvariant();
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
                        CurrentRequest.StartDate = parsedDate.Date; // Ensure only date part is stored
                        isValid = true;
                    }
                    // Attempt to parse a single day number
                    else if (int.TryParse(cleanedInput, out parsedNumber) && parsedNumber >= 1 && parsedNumber <= 31)
                    {
                        CurrentRequest.TempDay = parsedNumber;
                        CurrentState = ConversationState.AskForStartMonth;
                        errorMessage = "NEXT_STEP_OK"; // Signal successful temporary state transition
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
                            // Validate and set the complete date
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
                        parsedDate = parsedDate.Date;
                        if (parsedDate < CurrentRequest.StartDate.Date)
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
                            if (parsedDate.Date < CurrentRequest.StartDate.Date)
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

                case ConversationState.AskForDataScope:
                    string normalized = NormalizeCategory(cleanedInput);
                    if (!string.IsNullOrEmpty(normalized))
                    {
                        // User specified a canonical report category (e.g., Admissions)
                        CurrentRequest.DataScope = normalized;
                        CurrentRequest.Categories.Clear();
                        CurrentRequest.Categories.Add(normalized);
                        isValid = true;
                    }
                    else
                    {
                        // Check for Ward name match
                        if (WardList.Any(w => string.Equals(w, cleanedInput, StringComparison.OrdinalIgnoreCase)))
                        {
                            CurrentRequest.DataAction = WardList.First(w => string.Equals(w, cleanedInput, StringComparison.OrdinalIgnoreCase));
                            CurrentRequest.DataScope = "WARD_FILTER"; // Indicate it's a ward filter
                            isValid = true;
                        }
                        else
                        {
                            string low = cleanedInput.ToLowerInvariant();
                            if (low == "all" || low == "everything" || low == "any")
                            {
                                CurrentRequest.DataScope = "ALL"; // Report on all canonical categories
                                CurrentRequest.Categories.Clear(); // Ensure categories list is cleared if 'ALL' is chosen
                                isValid = true;
                            }
                            // Check for user filter
                            else if (low.StartsWith("user:"))
                            {
                                CurrentRequest.DataAction = cleanedInput; // e.g. "user:john"
                                CurrentRequest.DataScope = "USER_FILTER"; // Indicate it's a user filter
                                isValid = true;
                            }
                            else
                            {
                                errorMessage = "I didn't understand the scope. Try 'admissions', 'deaths', 'transfers', 'discharges', 'consultations', a ward name, 'user:username' or 'all'.";
                            }
                        }
                    }
                    break;

                case ConversationState.AskForSaveFormat:
                    string fmt = cleanedInput.Trim().ToLowerInvariant();
                    if (fmt == "txt" || fmt == "docx" || fmt == "xlsx")
                    {
                        CurrentRequest.FileFormat = fmt;
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Invalid format. Acceptable: txt, docx, xlsx.";
                    }
                    break;

                case ConversationState.AskForReportToOpen:
                    if (!string.IsNullOrWhiteSpace(cleanedInput))
                    {
                        CurrentRequest.ReportName = cleanedInput;
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please provide the report file name to open.";
                    }
                    break;

                case ConversationState.ConfirmationAndGenerate:
                    if (cleanedInput.Trim().Equals("yes", StringComparison.OrdinalIgnoreCase))
                    {
                        isValid = true;
                    }
                    else
                    {
                        errorMessage = "Please type 'Yes' to confirm generation.";
                    }
                    break;

                default:
                    isValid = true;
                    break;
            }

            return isValid;
        }

        // ----- NLP Normalizer: maps many phrasings to canonical categories used in DB -----
        /// <summary>
        /// Converts free text input into a standardized canonical report category if a match is found.
        /// </summary>
        private string NormalizeCategory(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            string s = input.Trim().ToLowerInvariant();

            if (s.Contains("admit") || s.Contains("admission") || s.Contains("new patient") || s.Contains("came in") || s.Contains("admissions"))
                return CatAdmission;

            if (s.Contains("death") || s.Contains("died") || s.Contains("passed away") || s.Contains("dead") || s.Contains("mortality") || s.Contains("mortuary"))
                return CatDeath;

            if (s.Contains("transfer") || s.Contains("moved") || s.Contains("inter ward") || s.Contains("interward") || s.Contains("transferred"))
            {
                // Prioritize 'out' or 'from' over 'in' or 'to' when both are present or ambiguous
                if (s.Contains("out") || s.Contains("from "))
                    return CatTransferOut;
                if (s.Contains("in") || s.Contains("into") || s.Contains("to "))
                    return CatTransferIn;
                return CatTransferIn; // Default to Transfer In if only 'transfer' is mentioned
            }

            if (s.Contains("discharge") || s.Contains("discharged") || s.Contains("release") || s.Contains("sent home"))
                return CatDischarge;

            if (s.Contains("consult") || s.Contains("clinic") || s.Contains("consultations") || s.Contains("visit"))
                return CatConsultations;

            // exact matches:
            if (string.Equals(s, "admission", StringComparison.OrdinalIgnoreCase)) return CatAdmission;
            if (string.Equals(s, "death", StringComparison.OrdinalIgnoreCase)) return CatDeath;
            if (string.Equals(s, "interwardtransferin", StringComparison.OrdinalIgnoreCase)) return CatTransferIn;
            if (string.Equals(s, "interwardtransferout", StringComparison.OrdinalIgnoreCase)) return CatTransferOut;
            if (string.Equals(s, "discharge", StringComparison.OrdinalIgnoreCase)) return CatDischarge;
            if (string.Equals(s, "consultations", StringComparison.OrdinalIgnoreCase)) return CatConsultations;

            return null;
        }

        // ----- Date parsing (robust) -----
        /// <summary>
        /// Attempts to parse a single date string using multiple robust formats.
        /// </summary>
        private bool TryParseSingleDate(string input, out DateTime date, out string error)
        {
            date = DateTime.MinValue;
            error = null;

            if (string.IsNullOrWhiteSpace(input))
            {
                error = "Date input cannot be empty.";
                return false;
            }

            string original = input.Trim();
            string normalized = original.Replace('\\', '-')
                                        .Replace('/', '-')
                                        .Replace('.', '-')
                                        .Replace('_', '-')
                                        .Trim();

            string[] formats = new string[]
            {
                "yyyy-MM-dd", "yyyy/MM/dd", // ISO format
                "dd-MM-yyyy", "dd/MM/yyyy", // British/European format
                "MM-dd-yyyy", "MM/dd/yyyy", // American format
                "d-M-yyyy",   "d/M/yyyy",   // Shorter British
                "M-d-yyyy",   "M/d/yyyy",   // Shorter American
                "yyyy-M-d",
                "dd-MMM-yyyy",
                "yyyyMMdd",
                "ddMMyyyy"
            };

            // 1. Try specific formats with invariant culture
            foreach (string fmt in formats)
            {
                string fmtNormalized = fmt.Replace('/', '-');
                if (DateTime.TryParseExact(normalized, fmtNormalized, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    return true;

                if (DateTime.TryParseExact(original, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    return true;
            }

            // 2. Try general parsing with common cultures
            if (DateTime.TryParse(normalized, new CultureInfo("en-GB"), DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(normalized, new CultureInfo("en-US"), DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(original, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                return true;

            if (DateTime.TryParse(original, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                return true;

            // 3. Try parsing without whitespace
            string compact = new string(original.Where(c => !char.IsWhiteSpace(c)).ToArray());
            if (DateTime.TryParse(compact, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                return true;

            error = "Invalid date. Please enter a valid date like 2025-10-01 or 01/10/2025.";
            return false;
        }

        // ----- Move to next conversation state -----
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
                case ConversationState.AskForDataScope:
                    // DataAction is now used for filtering within the scope, so we skip AskForDataAction
                    CurrentState = ConversationState.AskForSaveFormat;
                    break;
                case ConversationState.AskForSaveFormat:
                    CurrentState = ConversationState.ConfirmationAndGenerate;
                    break;
                case ConversationState.AskForReportToOpen:
                    CurrentState = ConversationState.Complete;
                    break;
            }
        }

        // ----- Database helpers (calls OleDb) -----
        /// <summary>
        /// Returns count of movements for given category and optional ward filter within a date range.
        /// </summary>
        /// <param name="startDateInclusive">The start date (inclusive).</param>
        /// <param name="endDateExclusive">The end date (exclusive).</param>
        /// <param name="category">The movement category (e.g., Admission).</param>
        /// <param name="wardFilter">The ward name to filter on (optional, defaults to ALL).</param>
        /// <returns>The count, or -1 on database error.</returns>
        public int GetPatientMovementCount(DateTime startDateInclusive, DateTime endDateExclusive, string category, string wardFilter = "ALL")
        {
            InitializeDatabase();
            int count = -1;
            string finalWardFilter = string.IsNullOrWhiteSpace(wardFilter) ? "ALL" : wardFilter;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(ConnectionString))
                {
                    conn.Open();

                    string sql = null;
                    OleDbCommand cmd = null;

                    // Base SQL for category and date range
                    string baseSql = "SELECT COUNT(*) FROM tblPatientMovement WHERE category = ? AND MovementDateTime >= ? AND MovementDateTime < ?";

                    if (category == CatConsultations)
                    {
                        // Note: Assuming Consultations are tracked in tblPatientMovement for simplicity
                        // or in tblPatientMaster (commented out the try/catch logic to simplify dependency on master table schema)
                        baseSql = "SELECT COUNT(*) FROM tblPatientMovement WHERE category = ? AND MovementDateTime >= ? AND MovementDateTime < ?";
                        cmd = new OleDbCommand(baseSql, conn);
                        cmd.Parameters.AddWithValue("?", category);
                        cmd.Parameters.AddWithValue("?", startDateInclusive);
                        cmd.Parameters.AddWithValue("?", endDateExclusive);
                    }
                    else if (category == CatAdmission || category == CatTransferIn)
                    {
                        if (finalWardFilter == "ALL")
                        {
                            cmd = new OleDbCommand(baseSql, conn);
                            cmd.Parameters.AddWithValue("?", category);
                            cmd.Parameters.AddWithValue("?", startDateInclusive);
                            cmd.Parameters.AddWithValue("?", endDateExclusive);
                        }
                        else
                        {
                            sql = "SELECT COUNT(*) FROM tblPatientMovement WHERE category = ? AND toWard = ? AND MovementDateTime >= ? AND MovementDateTime < ?";
                            cmd = new OleDbCommand(sql, conn);
                            cmd.Parameters.AddWithValue("?", category);
                            cmd.Parameters.AddWithValue("?", finalWardFilter);
                            cmd.Parameters.AddWithValue("?", startDateInclusive);
                            cmd.Parameters.AddWithValue("?", endDateExclusive);
                        }
                    }
                    else if (category == CatTransferOut || category == CatDischarge)
                    {
                        if (finalWardFilter == "ALL")
                        {
                            cmd = new OleDbCommand(baseSql, conn);
                            cmd.Parameters.AddWithValue("?", category);
                            cmd.Parameters.AddWithValue("?", startDateInclusive);
                            cmd.Parameters.AddWithValue("?", endDateExclusive);
                        }
                        else
                        {
                            sql = "SELECT COUNT(*) FROM tblPatientMovement WHERE category = ? AND fromWard = ? AND MovementDateTime >= ? AND MovementDateTime < ?";
                            cmd = new OleDbCommand(sql, conn);
                            cmd.Parameters.AddWithValue("?", category);
                            cmd.Parameters.AddWithValue("?", finalWardFilter);
                            cmd.Parameters.AddWithValue("?", startDateInclusive);
                            cmd.Parameters.AddWithValue("?", endDateExclusive);
                        }
                    }
                    else if (category == CatDeath)
                    {
                        // Death is generally tracked regardless of ward, but we include category and date filter
                        cmd = new OleDbCommand(baseSql, conn);
                        cmd.Parameters.AddWithValue("?", category);
                        cmd.Parameters.AddWithValue("?", startDateInclusive);
                        cmd.Parameters.AddWithValue("?", endDateExclusive);
                    }
                    
                    if (cmd != null)
                    {
                        object o = cmd.ExecuteScalar();
                        if (o != null && o != DBNull.Value)
                        {
                            count = Convert.ToInt32(o);
                        }
                        else
                        {
                            count = 0;
                        }
                    }
                }
            }
            catch
            {
                // return -1 on error
                count = -1;
            }

            return count;
        }

        /// <summary>
        /// Retrieves detailed movement records entered by a specified user within the date range.
        /// </summary>
        public List<Dictionary<string, object>> GetMovementsEnteredBy(string enteredBy, DateTime startDateInclusive, DateTime endDateExclusive)
        {
            InitializeDatabase();
            var results = new List<Dictionary<string, object>>();

            try
            {
                using (OleDbConnection conn = new OleDbConnection(ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT hospitalNumber, name, surname, gender, MovementDateTime, toWard, fromWard, category, enteredBy, movementID " +
                                 "FROM tblPatientMovement " +
                                 "WHERE enteredBy = ? AND MovementDateTime >= ? AND MovementDateTime < ? " +
                                 "ORDER BY MovementDateTime";

                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("?", enteredBy);
                        cmd.Parameters.AddWithValue("?", startDateInclusive);
                        cmd.Parameters.AddWithValue("?", endDateExclusive);

                        using (OleDbDataReader r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                var row = new Dictionary<string, object>();
                                for (int i = 0; i < r.FieldCount; i++)
                                {
                                    row[r.GetName(i)] = r.IsDBNull(i) ? null : r.GetValue(i);
                                }
                                results.Add(row);
                            }
                        }
                    }
                }
            }
            catch
            {
                // return empty list on error
            }

            return results;
        }

        /// <summary>
        /// Resets the password for a user account based on their ID number.
        /// (This function is provided in the original code, though less relevant to report generation)
        /// </summary>
        public string ResetUserPasswordByIdNumber(string idNumber)
        {
            if (string.IsNullOrWhiteSpace(idNumber))
                return null;

            InitializeDatabase();
            string newPassword = GenerateRandomPassword(8);

            try
            {
                using (OleDbConnection conn = new OleDbConnection(ConnectionString))
                {
                    conn.Open();
                    string sql = "UPDATE userAccounts SET password = ? WHERE idNumber = ?";

                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("?", newPassword);
                        cmd.Parameters.AddWithValue("?", idNumber);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            return newPassword;
                        else
                            return null;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private string GenerateRandomPassword(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var sb = new StringBuilder();
            var rnd = new Random();
            for (int i = 0; i < length; i++)
            {
                sb.Append(chars[rnd.Next(chars.Length)]);
            }
            return sb.ToString();
        }

        // ----- Report generation (main entry) -----
        /// <summary>
        /// Executes the report logic based on the collected request data, generates the content,
        /// and saves it to a file.
        /// </summary>
        public string ExecuteReportLogic()
        {
            InitializeDatabase();

            if (CurrentRequest.PrimaryAction == "OpenExisting")
            {
                // In a real application, you would launch the file here.
                return "Action Complete: Attempting to open existing report file: " + CurrentRequest.ReportName + "\nPath: " + Path.Combine(ReportDirectoryPath ?? "Unknown", CurrentRequest.ReportName);
            }

            StringBuilder report = new StringBuilder();
            report.AppendLine("Marondera Provincial Hospital - Movement Report");
            report.AppendLine("======================================================================");
            report.AppendLine("REPORT GENERATED ON: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            report.AppendLine("PERIOD: " + CurrentRequest.StartDateString + " to " + CurrentRequest.EndDateString);

            DateTime startInclusive = CurrentRequest.StartDate.Date;
            DateTime endExclusive = CurrentRequest.EndDate.Date.AddDays(1); // For >= start and < end logic

            string filter = (string.IsNullOrWhiteSpace(CurrentRequest.DataAction) ? "None" : CurrentRequest.DataAction);

            // --- Section 1: Summary Counts ---
            report.AppendLine();
            report.AppendLine("--- SUMMARY COUNTS ---");
            report.AppendLine("SCOPE: " + CurrentRequest.DataScope);
            report.AppendLine("FILTER: " + (CurrentRequest.DataScope == "USER_FILTER" ? CurrentRequest.DataAction : CurrentRequest.DataScope == "WARD_FILTER" ? CurrentRequest.DataAction : "None"));
            report.AppendLine("----------------------------------------------------------------------");

            // If DataScope is ALL or WARD_FILTER, show summary for all canonical categories
            if (string.Equals(CurrentRequest.DataScope, "ALL", StringComparison.OrdinalIgnoreCase) || CurrentRequest.DataScope == "WARD_FILTER" || string.IsNullOrWhiteSpace(CurrentRequest.DataScope))
            {
                var cats = new List<string> { CatAdmission, CatDeath, CatTransferIn, CatTransferOut, CatDischarge, CatConsultations };
                
                // Determine the ward filter based on the DataScope
                string wardFilter = (CurrentRequest.DataScope == "WARD_FILTER") ? CurrentRequest.DataAction : "ALL";

                foreach (var c in cats)
                {
                    int cnt = GetPatientMovementCount(startInclusive, endExclusive, c, wardFilter);
                    string countStr = (cnt >= 0) ? cnt.ToString() : "DB Error";
                    report.AppendLine(string.Format("{0,-20}: {1}", c, countStr));
                }
            }
            else if (CurrentRequest.Categories.Any())
            {
                // Only a single specific category was requested
                foreach (var cat in CurrentRequest.Categories.Distinct())
                {
                    int cnt = GetPatientMovementCount(startInclusive, endExclusive, cat, "ALL");
                    string countStr = (cnt >= 0) ? cnt.ToString() : "DB Error";
                    report.AppendLine(string.Format("{0,-20}: {1}", cat, countStr));
                }
            }

            // --- Section 2: Detailed User Activity ---
            if (CurrentRequest.DataScope == "USER_FILTER" && CurrentRequest.DataAction.StartsWith("user:", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = CurrentRequest.DataAction.Split(new char[] { ':' }, 2);
                if (parts.Length == 2)
                {
                    string username = parts[1];
                    var rows = GetMovementsEnteredBy(username, startInclusive, endExclusive);
                    report.AppendLine();
                    report.AppendLine("--- DETAIL: MOVEMENTS ENTERED BY USER: " + username.ToUpper() + " ---");
                    report.AppendLine(string.Format("Total Entries: {0}", rows.Count));
                    report.AppendLine("----------------------------------------------------------------------");

                    // Header
                    report.AppendLine(string.Format("{0,-12}|{1,-25}|{2,-20}|{3,-15}|{4,-15}",
                        "Hosp. #", "Name", "Time", "From Ward", "To Ward"));
                    report.AppendLine(new string('-', 80));

                    foreach (var r in rows)
                    {
                        object hosp = r.ContainsKey("hospitalNumber") ? r["hospitalNumber"] : "";
                        object name = r.ContainsKey("name") ? r["name"] : "";
                        object sname = r.ContainsKey("surname") ? r["surname"] : "";
                        DateTime mdt = (r.ContainsKey("MovementDateTime") && r["MovementDateTime"] is DateTime) ? (DateTime)r["MovementDateTime"] : DateTime.MinValue;
                        object fromW = r.ContainsKey("fromWard") ? r["fromWard"] : "";
                        object toW = r.ContainsKey("toWard") ? r["toWard"] : "";
                        object cat = r.ContainsKey("category") ? r["category"] : "";

                        report.AppendLine(string.Format("{0,-12}|{1,-25}|{2,-20}|{3,-15}|{4,-15} ({5})",
                            hosp,
                            name + " " + sname,
                            mdt.ToString("yyyy-MM-dd HH:mm"),
                            fromW ?? "-",
                            toW ?? "-",
                            cat
                            ));
                    }
                }
            }

            // Save report to file
            string format = (CurrentRequest.FileFormat ?? "txt").ToLowerInvariant();
            string fileName = "Report_" + CurrentRequest.DataScope.Replace(":", "_") + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + format;
            string finalPath = Path.Combine(ReportDirectoryPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);

            try
            {
                // Simple saving, assuming docx/xlsx are just text/csv files for this context
                File.WriteAllText(finalPath, report.ToString(), Encoding.UTF8);

                StringBuilder actionMsg = new StringBuilder();
                actionMsg.AppendLine("Report generated and saved successfully!");
                actionMsg.AppendLine("File format: " + format.ToUpper());
                actionMsg.AppendLine("Path: " + finalPath);
                actionMsg.AppendLine();
                actionMsg.AppendLine("--- Report Content Summary (saved to file) ---");
                actionMsg.AppendLine(report.ToString());

                return actionMsg.ToString();
            }
            catch (Exception ex)
            {
                return "Failure saving report to file at " + finalPath + ": " + ex.Message;
            }
        }

        // ----- Small helper to save report content directly if needed (kept for compatibility) -----
        public bool SaveReportContent(ReportRequest request, string fullPath)
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
            catch
            {
                return false;
            }
        }

        // ----- Prompts to guide the conversation -----
        /// <summary>
        /// Provides the user with the next instruction based on the current conversation state.
        /// </summary>
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
                case ConversationState.AskForDataScope:
                    return "What data would you like? (admissions, deaths, transfers, discharges, consultations) or type a ward name, 'user:username' or 'all'";
                case ConversationState.AskForSaveFormat:
                    return "Which file format? (txt, docx, xlsx)";
                case ConversationState.AskForReportToOpen:
                    return "Please type the exact filename of the report you wish to open (e.g., Report_20250101_123456.txt).";
                case ConversationState.ConfirmationAndGenerate:
                    return "Ready to generate report for" +"CurrentRequest.StartDateString"+ "to" +"CurrentRequest.EndDateString"+", Scope: "+"CurrentRequest.DataScope"+", Format: "+"CurrentRequest.FileFormat"+".\nType 'Yes' to confirm and generate the report.";
                default:
                    return "The process is complete or awaiting next step.";
            }
        }
    }
}