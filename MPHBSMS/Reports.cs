using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Data.OleDb;
using Microsoft.Reporting.WinForms;


// Ensure this matches your project's namespace
namespace MPHBSMS
{
    public partial class Reports : Form
    {
        // 1. DECLARATION: The ReportManager instance
        private ReportManager _manager;

        public Reports()
        {
            InitializeComponent();

        }

        // ➡️ 1. INITIALIZATION: Form Load Event
        private void Reports_Load(object sender, EventArgs e)
        {

            comboBox2.Items.AddRange(new string[]
    {
        "Mental Health Unit",
        "Female Ward",
        "Paedatric Ward",
        "Male Ward",
        "PostNatal Ward",
        "NeoNatal Ward",
        "AnteNatal Ward",
        "Labor Ward",
        "Accident and Emergence",
        "ALL"
    });
            comboBox2.SelectedIndex = 0;

            comboBox1.Items.AddRange(new string[]
    {
        "Admission",
        "InterWardTransferIn",
        "Discharge",
        "InterWardTransferOut",
        "Death",
        "ALL"
    });
            comboBox1.SelectedIndex = 0;

            if (_manager == null)
            {
                _manager = new ReportManager();
            }

            DisplayBotGreeting();
            UpdateUI(initialCall: true);

            // Set the state immediately to prevent the first prompt from being processed as input
            _manager.CurrentState = ConversationState.AskForPrimaryAction;

            // Focus the richTextBox so the user can start typing right away
            richTextBox1.Focus();
        }

        // ➡️ 2. INPUT HANDLER: Captures Enter key in richTextBox1
        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                // 1. Get the last line of text (the command)
                string userInput = GetLastLineInput();

                if (string.IsNullOrWhiteSpace(userInput))
                {
                    return;
                }

                // Clear the input line *BEFORE* logging and processing.
                ClearLastLineInput();

                // 2. Log the user's input permanently (Color: White)
                ChatLogger.AppendText(richTextBox1, "You: " + userInput, Color.White, true);

                // 3. Call the manager to process the command
                ProcessResult result = _manager.ProcessAndAdvance(userInput);

                if (!result.Success)
                {
                    // If validation failed, log the error message (Color: Red)
                    ChatLogger.AppendText(richTextBox1, "MPH Bot Error: " + result.ErrorMessage, Color.Red);
                }
                else
                {
                    // Log the successful state transition
                    ChatLogger.AppendText(richTextBox1, "--- State Advanced to: " + _manager.CurrentState.ToString() + " ---", Color.OrangeRed, true);
                }

                // 4. Update UI (displays the next prompt or executes the report)
                UpdateUI();
            }
        }

        // 🔄 3. RESET BUTTON LOGIC (Assuming button4 is the Reset button)
        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            _manager = new ReportManager();
            DisplayBotGreeting();
            UpdateUI();
        }

        // ❌ 4. REDUNDANT BUTTON (button5)
        private void button5_Click(object sender, EventArgs e)
        {
            // Simply call the reset logic
            button4_Click(sender, e);
        }

        // 💡 5. CORE UI UPDATER
        private void UpdateUI(bool initialCall = false)
        {
            if (_manager.CurrentState == ConversationState.ConfirmationAndGenerate)
            {
                string finalLog = _manager.ExecuteReportLogic();

                ChatLogger.AppendText(richTextBox1, "\n--- EXECUTION REPORT ---", Color.Gray, true);
                ChatLogger.AppendText(richTextBox1, finalLog, Color.White, false);
                ChatLogger.AppendText(richTextBox1, "------------------------\n", Color.Gray, true);

                _manager.CurrentState = ConversationState.Complete;

                UpdateUI();
                return;
            }

            string prompt = _manager.GetCurrentPrompt();

            if (_manager.CurrentState != ConversationState.Greeting)
            {
                // Bot prompt (Color: White, Prefix: MPH Bot)
                ChatLogger.AppendText(richTextBox1, "MPH Bot: " + prompt, Color.White, false);
            }
        }

        // 💡 6. INITIAL GREETING 
        private void DisplayBotGreeting()
        {
            string greeting = "Welcome to the Report Generator! I'm here to guide you through creating or opening a report.";
            // Bot greeting (Color: White, Prefix: MPH Bot)
            ChatLogger.AppendText(richTextBox1, "MPH Bot: " + greeting, Color.White, true);
        }

        // 💡 7. HELPER METHODS FOR RICH TEXT BOX INPUT
        private string GetLastLineInput()
        {
            string fullText = richTextBox1.Text;
            int lastNewLine = fullText.LastIndexOf(Environment.NewLine);

            if (lastNewLine > 0)
            {
                return fullText.Substring(lastNewLine + Environment.NewLine.Length).Trim();
            }
            return fullText.Trim();
        }

        private void ClearLastLineInput()
        {
            int lastNewLine = richTextBox1.Text.LastIndexOf(Environment.NewLine);

            if (lastNewLine > 0)
            {
                richTextBox1.Select(lastNewLine + Environment.NewLine.Length, richTextBox1.Text.Length - (lastNewLine + Environment.NewLine.Length));
                richTextBox1.SelectedText = string.Empty;
            }
            else
            {
                richTextBox1.Clear();
                DisplayBotGreeting();
            }
        }

        // === Navigation/Placeholder Event Handlers ===
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 obj = new Form1();
            this.Close();
            obj.Show();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Menu obj = new Menu();
            this.Close();
            obj.Show();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Visual obj = new Visual();
            this.Close();
            obj.Show();
        }
        private void laborWardToolStripMenuItem_Click(object sender, EventArgs e) { /* Menu logic */ }
        private void richTextBox1_TextChanged(object sender, EventArgs e) { /* Empty */ }

        private void button5_Click_1(object sender, EventArgs e)
        {
            // --- NEW: Initialize Database before connection ---
            DatabaseHelper.InitializeDatabase();
            /*
            try
            {
                // 1. Get Selections (Ward only, we will ignore Category dropdown for the summary labels)
                string selectedWard = "ALL";
                if (comboBox2.SelectedItem != null)
                {
                    var tmp = comboBox2.SelectedItem.ToString().Trim();
                    if (!string.IsNullOrEmpty(tmp)) selectedWard = tmp;
                }

                // 2. Validate input
                if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    MessageBox.Show("Please enter both a valid Start Date and End Date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                // 3. Parse Dates
                DateTime startDate;
                if (!DateTime.TryParse(textBox1.Text.Trim(), out startDate))
                {
                    MessageBox.Show("Start Date is not a valid date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                DateTime endDate;
                if (!DateTime.TryParse(textBox2.Text.Trim(), out endDate))
                {
                    MessageBox.Show("End Date is not a valid date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                startDate = startDate.Date;
                DateTime exclusiveEndDate = endDate.Date.AddDays(1);

                // 4. Define the categories to count and their target labels
                // KEY = Database Value, VALUE = Label to update
                // CHECK SPELLING: Ensure the Keys match your database 'category' column exactly!
                var categories = new Dictionary<string, Label>
    {
        { "Admission", label5 },
        { "InterWardTransferIn", label6 },
        { "Discharge", label7 },
        { "InterWardTransferOut", label8 },
        { "Death", label9 }
    };

                // 5. Execute Loop
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    foreach (var kvp in categories)
                    {
                        string categoryToCount = kvp.Key;
                        Label targetLabel = kvp.Value;

                        // Build Query
                        string query =
                            @"SELECT COUNT(*) 
                  FROM ( 
                      SELECT DISTINCT TM.hospitalNumber 
                      FROM tblPatientMaster AS PM 
                      INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber ";

                        string whereClause = "WHERE 1=1 ";

                        // 1. Filter by the specific category for this label
                        whereClause += "AND TM.category = ? ";

                        // 2. Filter by Ward (if selected)
                        bool useWard = !string.Equals(selectedWard, "ALL", StringComparison.OrdinalIgnoreCase);
                        if (useWard)
                            whereClause += "AND PM.currentWard = ? ";

                        // 3. Filter by Date
                        whereClause += "AND TM.MovementDateTime >= ? AND TM.MovementDateTime < ? ";

                        query += whereClause + ") AS T";

                        using (OleDbCommand cmd = new OleDbCommand(query, con))
                        {
                            // Parameters must be added in the exact order they appear in the query (?)

                            // 1. Category Parameter
                            cmd.Parameters.AddWithValue("?", categoryToCount);

                            // 2. Ward Parameter
                            if (useWard)
                                cmd.Parameters.AddWithValue("?", selectedWard);

                            // 3. Date Parameters
                            cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            // Execute
                            object result = cmd.ExecuteScalar();
                            int totalCount = 0;

                            if (result != null && result != DBNull.Value)
                                totalCount = Convert.ToInt32(result);

                            // Update the specific label
                            targetLabel.Text = totalCount.ToString("D2");
                        }
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error occurred\n" + error.Message,
                                "Marondera Provincial Hospital",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }*/
    /*    try
{
    // ============================
    // LOG FILE PATH
    // ============================
    string logPath = @"C:\MPH_Logs\QueryLog.txt";

    // Ensure folder exists
    if (!Directory.Exists(@"C:\MPH_Logs"))
        Directory.CreateDirectory(@"C:\MPH_Logs");


    // 1. Get Selections (null-safe)
    string criteria = "ALL";
    if (comboBox1.SelectedItem != null)
    {
        var tmp = comboBox1.SelectedItem.ToString().Trim();
        if (!string.IsNullOrEmpty(tmp)) criteria = tmp;
    }

    string selectedWard = "ALL";
    if (comboBox2.SelectedItem != null)
    {
        var tmp = comboBox2.SelectedItem.ToString().Trim();
        if (!string.IsNullOrEmpty(tmp)) selectedWard = tmp;
    }

    // 2. Validate input
    if (string.IsNullOrWhiteSpace(textBox1.Text) ||
        string.IsNullOrWhiteSpace(textBox2.Text))
    {
        MessageBox.Show("Please enter both a valid Start Date and End Date.",
                        "Marondera Provincial Hospital",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
        return;
    }

    // 3. Parse Dates
    DateTime startDate;
    if (!DateTime.TryParse(textBox1.Text.Trim(), out startDate))
    {
        MessageBox.Show("Start Date is not a valid date.",
                        "Marondera Provincial Hospital",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
        return;
    }

    DateTime endDate;
    if (!DateTime.TryParse(textBox2.Text.Trim(), out endDate))
    {
        MessageBox.Show("End Date is not a valid date.",
                        "Marondera Provincial Hospital",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
        return;
    }

    startDate = startDate.Date;
    DateTime exclusiveEndDate = endDate.Date.AddDays(1);

    // 4. Build Query (WITH DISTINCT)
    string query = @"SELECT COUNT(HospitalID) AS TotalCount 
        FROM (
                SELECT DISTINCT TM.hospitalNumber AS HospitalID
                FROM tblPatientMaster AS PM
                INNER JOIN tblPatientMovement AS TM
                ON PM.hospitalNumber = TM.hospitalNumber )";

    
            
    string whereClause = "WHERE 1=1 ";

    bool useCategory = !string.Equals(criteria, "ALL", StringComparison.OrdinalIgnoreCase);
    bool useWard = !string.Equals(selectedWard, "ALL", StringComparison.OrdinalIgnoreCase);

    if (useCategory)
        whereClause += "AND TM.category = ? ";

    if (useWard)
        whereClause += "AND PM.currentWard = ? ";

    // CORRECT date column
    whereClause += "AND TM.MovementDateTime >= ? AND TM.MovementDateTime < ? ";

    query += whereClause;

    // 5. Execute SQL
    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    using (OleDbCommand cmd = new OleDbCommand(query, con))
    {
        // Order must match '?' sequence
        if (useCategory)
            cmd.Parameters.AddWithValue("?", criteria);

        if (useWard)
            cmd.Parameters.AddWithValue("?", selectedWard);

        cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
        cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

        // ============================
        // LOG SQL + PARAMETERS
        // ============================
        using (StreamWriter sw = new StreamWriter(logPath, true))
        {
            sw.WriteLine("======================================================");
            sw.WriteLine("DATE: " + DateTime.Now);
            sw.WriteLine("EXECUTED QUERY:");
            sw.WriteLine(query);
            sw.WriteLine("\nParameters:");

            foreach (OleDbParameter p in cmd.Parameters)
                sw.WriteLine(" - " + p.Value);

            sw.WriteLine("======================================================\n");
        }

        // Reset UI values
        label5.Text = "00";

        con.Open();

        object result = cmd.ExecuteScalar();
        int totalCount = 0;

        if (result != null && result != DBNull.Value)
            totalCount = Convert.ToInt32(result);

        label5.Text = totalCount.ToString("D2");
    }
}
catch (Exception error)
{
    MessageBox.Show("Error occurred\n" + error.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
}*/

            try
            {
                // 1. Get Selections (null-safe)
                string criteria = "ALL";
                if (comboBox1.SelectedItem != null)
                {
                    var tmp = comboBox1.SelectedItem.ToString().Trim();
                    if (!string.IsNullOrEmpty(tmp)) criteria = tmp;
                }

                string selectedWard = "ALL";
                if (comboBox2.SelectedItem != null)
                {
                    var tmp = comboBox2.SelectedItem.ToString().Trim();
                    if (!string.IsNullOrEmpty(tmp)) selectedWard = tmp;
                }

                // 2. Validate input
                if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    MessageBox.Show("Please enter both a valid Start Date and End Date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                // 3. Parse Dates
                DateTime startDate;
                if (!DateTime.TryParse(textBox1.Text.Trim(), out startDate))
                {
                    MessageBox.Show("Start Date is not a valid date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                DateTime endDate;
                if (!DateTime.TryParse(textBox2.Text.Trim(), out endDate))
                {
                    MessageBox.Show("End Date is not a valid date.",
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                startDate = startDate.Date;
                DateTime exclusiveEndDate = endDate.Date.AddDays(1);

                // 4. Build Query (CORRECTED STRUCTURE)
                // We open the subquery, but we DO NOT close it with ')' yet.
                string query =
                    @"SELECT COUNT(*) AS TotalCount 
          FROM ( 
              SELECT DISTINCT TM.hospitalNumber 
              FROM tblPatientMaster AS PM 
              INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber ";

                // We start the WHERE clause inside the subquery
                string whereClause = "WHERE 1=1 ";

                bool useCategory = !string.Equals(criteria, "ALL", StringComparison.OrdinalIgnoreCase);
                bool useWard = !string.Equals(selectedWard, "ALL", StringComparison.OrdinalIgnoreCase);

                if (useCategory)
                    whereClause += "AND TM.category = ? ";

                if (useWard)
                    whereClause += "AND PM.currentWard = ? ";

                // Date filtering on the Movement table
                whereClause += "AND TM.MovementDateTime >= ? AND TM.MovementDateTime < ? ";

                // Combine the parts
                query += whereClause;

                // NOW we close the subquery parentheses and give it an alias [T]
                query += ") AS T";

                // 5. Execute SQL
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                using (OleDbCommand cmd = new OleDbCommand(query, con))
                {
                    // Must match '?' order exactly
                    if (useCategory)
                        cmd.Parameters.AddWithValue("?", criteria);

                    if (useWard)
                        cmd.Parameters.AddWithValue("?", selectedWard);

                    // Access/OleDb cares strictly about the order of parameters
                    cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                    cmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                    label5.Text = "00"; // reset labels

                    con.Open();

                    object result = cmd.ExecuteScalar();
                    int totalCount = 0;

                    if (result != null && result != DBNull.Value)
                        totalCount = Convert.ToInt32(result);

                    label5.Text = totalCount.ToString("D2");
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error occurred\n" + error.Message,
                                "Marondera Provincial Hospital",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }
    }
}