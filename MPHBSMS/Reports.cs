using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Data.OleDb;


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

try
{
    // 1. Get Selections
    string criteria = comboBox1.SelectedItem.ToString();
    string selectedWard = comboBox2.SelectedItem.ToString(); // Get selected ward from new ComboBox

    // 2. Parse Dates
    DateTime startDate = Convert.ToDateTime(textBox1.Text).Date; 
    // Set end date to end of day (23:59:59)
    DateTime endDate = Convert.ToDateTime(textBox2.Text).Date.AddDays(1).AddSeconds(-1);

    // 3. Dynamic Query Construction
    string query = "SELECT COUNT(TM.hospitalNumber) AS TotalCount " +
                   "FROM tblPatientMaster AS PM " +
                   "INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber ";
    
    // Start with a base WHERE clause (1=1 is useful for starting conditional concatenation)
    string whereClause = "WHERE 1=1 "; 

    // --- Ward and Category Filtering ---

    // Category Filter: Check if a specific category is selected
    if (criteria.ToUpper() != "ALL")
    {
        whereClause += "AND TM.category = ? ";
    }
    
    // Ward Filter: Check if a specific ward is selected
    if (selectedWard.ToUpper() != "ALL")
    {
        whereClause += "AND PM.currentWard = ? ";
    }
    
    // Date Filter (Always included)
    whereClause += "AND TM.movementDate BETWEEN ? AND ?";
    
    query += whereClause;

    // --- 4. Execute Command ---
    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        using (OleDbCommand cmd = new OleDbCommand(query, con))
        {
            // 5. Add Parameters (Order is CRITICAL for OLEDB '?' placeholders)
            
            // Category Parameter (1st if criteria is not "ALL")
            if (criteria.ToUpper() != "ALL")
            {
                cmd.Parameters.AddWithValue("@category", criteria);
            }

            // Ward Parameter (2nd if selectedWard is not "ALL")
            if (selectedWard.ToUpper() != "ALL")
            {
                cmd.Parameters.AddWithValue("@ward", selectedWard);
            }

            // Date Parameters (Always last)
            cmd.Parameters.AddWithValue("@startDate", startDate);
            cmd.Parameters.AddWithValue("@endDate", endDate);


            // Reset all dashboard labels to "00"
            label5.Text = "00";
            label6.Text = "00";
            label7.Text = "00";
            label8.Text = "00";
            label9.Text = "00";

            // Open Connection
            con.Open();
            
            // Execute Scalar returns the single value (the count)
            object result = cmd.ExecuteScalar();
            int totalCount = result != DBNull.Value ? Convert.ToInt32(result) : 0;
            string countText = totalCount.ToString("D2");
            
            // Assuming label5 is the main result display for a single-selection filter.
            label5.Text = countText;
        } // cmd is automatically Disposed
    } // con is automatically Closed and Disposed
}
catch (Exception error)
{
    // Error handling remains
    MessageBox.Show("Error occured\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
// The finally block is now unnecessary as the 'using' statement handles connection disposal automatically.
        }
    }
}