using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;

// Dummy NLP processor for demonstration. Replace with advanced NLP if needed.
public class NLPProcessor
{
    public string ExtractIntent(string message)
    {
        message = message.ToLower();
        if (message.Contains("admit") || message.Contains("patient") || message.Contains("who is"))
            return "PatientInfo";
        if (message.Contains("move") || message.Contains("transfer") || message.Contains("ward"))
            return "MovementInfo";
        if (message.Contains("report") || message.Contains("summary") || message.Contains("generate"))
            return "ReportManager";
        if (message.Contains("open") && message.Contains("report"))
            return "OpenReport";
        if (message.Contains("count") || message.Contains("how many"))
            return "CountInfo";
        return "Unknown";
    }

    public Dictionary<string, string> ExtractEntities(string message)
    {
        var entities = new Dictionary<string, string>();
        foreach (var word in message.Split(' '))
        {
            if (word.All(char.IsDigit) && word.Length > 4)
                entities["hospitalNumber"] = word;
        }
        return entities;
    }
}

public enum MessageSide { Left, Right }

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
    public string PrimaryAction { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string DataScope { get; set; }
    public string DataAction { get; set; }
    public string FileFormat { get; set; }
    public string ReportName { get; set; }
    public List<string> Categories { get; set; }

    public ReportRequest()
    {
        Categories = new List<string>();
    }
}

public class MphFullFeatureChatBot
{
    private OleDbConnection _con;
    private string _currentUser;
    private NLPProcessor _nlp;
    private List<Tuple<MessageSide, string>> _conversationHistory;
    private ConversationState _state;
    private ReportRequest _reportRequest;
    private readonly string[] ValidReportCategories = { "Admission", "Deaths", "Transfers", "Discharges", "Consultations" };
    private readonly string ReportDirectoryPath;
    private static readonly string[] AcceptedDateFormats = new string[]
    {
        "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy", "M/d/yyyy", "dd-MMM-yyyy"
    };

    public MphFullFeatureChatBot()
    {
        _nlp = new NLPProcessor();
        _conversationHistory = new List<Tuple<MessageSide, string>>();
        _state = ConversationState.Greeting;
        _reportRequest = new ReportRequest();

        // Always get the current user for every request
        _currentUser = MPHBSMS.CurrentUser;
        _con = new OleDbConnection(DatabaseHelper.ConnectionString);
        _con.Open();
        DatabaseHelper.InitializeDatabase();

        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        ReportDirectoryPath = Path.Combine(documentsPath, "MyReportGeneratorFiles");
        if (!Directory.Exists(ReportDirectoryPath)) Directory.CreateDirectory(ReportDirectoryPath);
    }

    public string HandleMessage(string userMessage)
    {
        _currentUser = MPHBSMS.CurrentUser;
        _conversationHistory.Add(new Tuple<MessageSide, string>(MessageSide.Right, "User (" + _currentUser + "): " + userMessage));

        string intent = _nlp.ExtractIntent(userMessage);
        Dictionary<string, string> entities = _nlp.ExtractEntities(userMessage);

        string response;

        switch (intent)
        {
            case "PatientInfo":
                response = RespondWithPatientInfo(entities);
                break;
            case "MovementInfo":
                response = RespondWithMovementInfo(entities);
                break;
            case "ReportManager":
                response = HandleReportManager(userMessage);
                break;
            case "OpenReport":
                response = HandleOpenReport(userMessage);
                break;
            case "CountInfo":
                response = RespondWithCounts(entities);
                break;
            default:
                response = "I'm here to assist with patients, movements, reports, and counts. How can I help?";
                break;
        }

        _conversationHistory.Add(new Tuple<MessageSide, string>(MessageSide.Left, "Bot: " + response));
        return response;
    }

    // Retrieve from BOTH PatientMaster and PatientMovement
    private string RespondWithPatientInfo(Dictionary<string, string> entities)
    {
        string hospitalNumber = entities.ContainsKey("hospitalNumber") ? entities["hospitalNumber"] : null;
        string masterQuery = "SELECT * FROM tblPatientMaster";
        if (hospitalNumber != null)
            masterQuery += " WHERE hospitalNumber = '" + hospitalNumber + "'";

        string info = "";

        using (OleDbCommand cmd = new OleDbCommand(masterQuery, _con))
        using (OleDbDataReader reader = cmd.ExecuteReader())
        {
            if (reader.Read())
            {
                info = "Patient Info:\n";
                info += "Name: " + reader["name"].ToString() + "\n";
                info += "Ward: " + reader["currentWard"].ToString() + "\n";
                info += "Admitted: " + reader["isAdmitted"].ToString() + "\n";
                info += "Admission Date: " + reader["admissionDate"].ToString() + "\n";
                info += "Discharge Date: " + reader["dischargeDate"].ToString() + "\n";
                info += "Gender: " + reader["gender"].ToString() + "\n";
            }
            else
            {
                return "No patient found for that number.";
            }
        }

        // Get movement info for the patient
        string movementQuery = "SELECT * FROM tblPatientMovement";
        if (hospitalNumber != null)
            movementQuery += " WHERE hospitalNumber = '" + hospitalNumber + "'";
        List<string> movements = new List<string>();
        using (OleDbCommand cmd = new OleDbCommand(movementQuery, _con))
        using (OleDbDataReader reader = cmd.ExecuteReader())
        {
            while (reader.Read())
            {
                string movement = "Movement: "
                    + reader["MovementDateTime"].ToString()
                    + " | " + reader["fromWard"].ToString()
                    + " → " + reader["toWard"].ToString()
                    + " (" + reader["category"].ToString() + ")";
                movements.Add(movement);
            }
        }

        if (movements.Count > 0)
        {
            info += "\nMovement History:\n" + string.Join("\n", movements);
        }
        else
        {
            info += "\nNo movement records found for this patient.";
        }

        return info;
    }

    private string RespondWithMovementInfo(Dictionary<string, string> entities)
    {
        string query = "SELECT * FROM tblPatientMovement WHERE enteredBy = '" + _currentUser + "'";
        if (entities.ContainsKey("hospitalNumber"))
            query += " AND hospitalNumber = '" + entities["hospitalNumber"] + "'";
        using (OleDbCommand cmd = new OleDbCommand(query, _con))
        using (OleDbDataReader reader = cmd.ExecuteReader())
        {
            List<string> movements = new List<string>();
            while (reader.Read())
            {
                movements.Add(reader["MovementDateTime"].ToString() + ": " +
                              reader["fromWard"].ToString() + " → " +
                              reader["toWard"].ToString() + " (" +
                              reader["category"].ToString() + ")");
            }
            return movements.Count > 0 ? string.Join("\n", movements) : "No movement records found.";
        }
    }

    // Count patients and movements, optionally by hospitalNumber
    private string RespondWithCounts(Dictionary<string, string> entities)
    {
        string hospitalNumber = entities.ContainsKey("hospitalNumber") ? entities["hospitalNumber"] : null;
        string info = "";

        // Count patients
        string patientCountQuery = "SELECT COUNT(*) FROM tblPatientMaster";
        if (hospitalNumber != null)
            patientCountQuery += " WHERE hospitalNumber = '" + hospitalNumber + "'";
        int patientCount = 0;
        using (OleDbCommand cmd = new OleDbCommand(patientCountQuery, _con))
        {
            object result = cmd.ExecuteScalar();
            patientCount = (result != null) ? Convert.ToInt32(result) : 0;
        }
        info = "Number of matching patients: " + patientCount.ToString();

        // Count movements
        string movementCountQuery = "SELECT COUNT(*) FROM tblPatientMovement";
        if (hospitalNumber != null)
            movementCountQuery += " WHERE hospitalNumber = '" + hospitalNumber + "'";
        int movementCount = 0;
        using (OleDbCommand cmd = new OleDbCommand(movementCountQuery, _con))
        {
            object result = cmd.ExecuteScalar();
            movementCount = (result != null) ? Convert.ToInt32(result) : 0;
        }
        info += "\nNumber of matching movements: " + movementCount.ToString();

        return info;
    }

    private string HandleReportManager(string userMessage)
    {
        switch (_state)
        {
            case ConversationState.Greeting:
                _state = ConversationState.AskForPrimaryAction;
                return "Do you want to generate a new report or open an existing one?";
            case ConversationState.AskForPrimaryAction:
                if (userMessage.ToLower().Contains("generate"))
                {
                    _reportRequest.PrimaryAction = "GenerateReport";
                    _state = ConversationState.AskForTimeFrame;
                    return "What is the time frame? (e.g., '2023-01-01 to 2023-10-31')";
                }
                else if (userMessage.ToLower().Contains("open"))
                {
                    _reportRequest.PrimaryAction = "OpenExisting";
                    _state = ConversationState.AskForReportToOpen;
                    return "What is the name of the report you want to open?";
                }
                else
                {
                    return "Please specify if you want to generate a report or open one.";
                }
            case ConversationState.AskForTimeFrame:
                DateTime start, end;
                string error;
                if (TryParseTimeFrame(userMessage, out start, out end, out error))
                {
                    _reportRequest.StartDate = start;
                    _reportRequest.EndDate = end;
                    _state = ConversationState.AskForDataScope;
                    return "Which categories? (Admission, Deaths, Transfers, Discharges, Consultations, or 'All')";
                }
                else
                {
                    return "Please enter a valid time frame (e.g., '2023-01-01 to 2023-10-31')";
                }
            case ConversationState.AskForDataScope:
                if (userMessage.ToLower().Contains("all"))
                {
                    _reportRequest.DataScope = "All Data";
                    _state = ConversationState.AskForDataAction;
                    return "Do you want to save, print, or display the report?";
                }
                else
                {
                    string[] cats = userMessage.Split(',');
                    foreach (string cat in cats)
                    {
                        string trimmedCat = cat.Trim();
                        if (ValidReportCategories.Any(c => c.Equals(trimmedCat, StringComparison.OrdinalIgnoreCase)))
                            _reportRequest.Categories.Add(trimmedCat);
                    }
                    if (_reportRequest.Categories.Count > 0)
                    {
                        _reportRequest.DataScope = "By Category";
                        _state = ConversationState.AskForDataAction;
                        return "Do you want to save, print, or display the report?";
                    }
                    else
                    {
                        return "Invalid categories. Valid: " + string.Join(", ", ValidReportCategories) + " or 'All'";
                    }
                }
            case ConversationState.AskForDataAction:
                if (userMessage.ToLower().Contains("save"))
                {
                    _reportRequest.DataAction = "Save File";
                    _state = ConversationState.AskForSaveFormat;
                    return "What format? (PDF or DOCX)";
                }
                else if (userMessage.ToLower().Contains("print"))
                {
                    _reportRequest.DataAction = "Print";
                    _state = ConversationState.ConfirmationAndGenerate;
                    return "Ready to print. Type 'Confirm' to proceed.";
                }
                else if (userMessage.ToLower().Contains("display"))
                {
                    _reportRequest.DataAction = "Display";
                    _state = ConversationState.ConfirmationAndGenerate;
                    return "Ready to display. Type 'Confirm' to proceed.";
                }
                else
                {
                    return "Please specify: save, print, or display.";
                }
            case ConversationState.AskForSaveFormat:
                if (userMessage.ToLower().Contains("pdf"))
                {
                    _reportRequest.FileFormat = "PDF";
                    _state = ConversationState.ConfirmationAndGenerate;
                    return "Ready to save as PDF. Type 'Confirm' to proceed.";
                }
                else if (userMessage.ToLower().Contains("docx"))
                {
                    _reportRequest.FileFormat = "DOCX";
                    _state = ConversationState.ConfirmationAndGenerate;
                    return "Ready to save as DOCX. Type 'Confirm' to proceed.";
                }
                else
                {
                    return "Please choose PDF or DOCX.";
                }
            case ConversationState.ConfirmationAndGenerate:
                if (userMessage.ToLower().Contains("confirm") || userMessage.ToLower().Contains("yes"))
                {
                    _state = ConversationState.Complete;
                    return ExecuteReportLogic();
                }
                else
                {
                    return "Please type 'Confirm' to proceed.";
                }
            case ConversationState.Complete:
                _state = ConversationState.Greeting;
                _reportRequest = new ReportRequest();
                return "Report complete. How else may I help?";
            default:
                return "I'm ready to help with reports.";
        }
    }

    private string HandleOpenReport(string userMessage)
    {
        string fileName = userMessage.Trim();
        string fullPath = Path.Combine(ReportDirectoryPath, fileName);
        if (File.Exists(fullPath))
            return "Opened report: " + fileName;
        if (File.Exists(fullPath + ".pdf"))
            return "Opened report: " + fileName + ".pdf";
        if (File.Exists(fullPath + ".docx"))
            return "Opened report: " + fileName + ".docx";
        return "Report not found.";
    }

    private string ExecuteReportLogic()
    {
        if (_reportRequest.PrimaryAction == "OpenExisting")
        {
            return "Action Complete: Attempting to open existing report file: " + _reportRequest.ReportName;
        }
        string fileName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "." + ((_reportRequest.FileFormat != null) ? _reportRequest.FileFormat.ToLower() : "txt");
        string fullFilePath = Path.Combine(ReportDirectoryPath, fileName);

        string content = "Report Generated: " + DateTime.Now.ToString() + "\n" +
                         "User: " + _currentUser + "\n" +
                         "Time Frame: " + _reportRequest.StartDate.ToString("yyyy-MM-dd") + " to " + _reportRequest.EndDate.ToString("yyyy-MM-dd") + "\n" +
                         "Scope: " + _reportRequest.DataScope + " (Categories: " + string.Join(", ", _reportRequest.Categories) + ")\n" +
                         "Action: Save as " + _reportRequest.FileFormat;

        File.WriteAllText(fullFilePath, content);
        return "Success: Report saved as " + _reportRequest.FileFormat + " to " + fullFilePath;
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

    public List<Tuple<MessageSide, string>> GetConversationHistory()
    {
        return _conversationHistory;
    }

    ~MphFullFeatureChatBot()
    {
        if (_con != null && _con.State == System.Data.ConnectionState.Open)
            _con.Close();
    }
}