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
        "Accident and Emergence"
    });
            comboBox2.SelectedIndex = 0;

            comboBox1.Items.AddRange(new string[]
    {
        "Admission",
        "InterWardTransferIn",
        "Discharge",
        "InterWardTransferOut",
        "Death"
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
            
            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                con.Open();
                try
                {
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
                    string category = comboBox1.SelectedItem.ToString();
                    string selectedWard = comboBox2.SelectedItem.ToString();

                    //Admissions only
                    if (category == "Admission")
                    {
                        string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ? AND TM.ToWard = ? AND PM.admissionDateTime >= ? AND PM.admissionDateTime < ?
                               )";



                        using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                        {
                            wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            int count = (int)wardAdmissionsCmd.ExecuteScalar();

                                label5.Text = count.ToString();
                          }
                    }

                    //Inter Ward Transfer In only
                    else if (category == "InterWardTransferIn")
                    {
                        string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ? AND TM.ToWard = ? AND TM.MovementDateTime >= ? AND TM.MovementDateTime < ?
                               )";



                        using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                        {
                            wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            int count = (int)wardAdmissionsCmd.ExecuteScalar();

                                label6.Text = count.ToString();
                            }
                    }

                    //Disharges only
                    else if (category == "Discharge")
                    {
                        string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ? AND TM.ToWard = ? AND PM.dischargeDateTime >= ? AND PM.dischargeDateTime < ?
                               )";



                        using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                        {
                            wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            int count = (int)wardAdmissionsCmd.ExecuteScalar();

                                label7.Text = count.ToString();
                          }
                    }

                    //Inter Ward Transfer Out only
                    else if (category == "InterWardTransferOut")
                    {
                        string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ? AND TM.ToWard = ? AND TM.MovementDateTime >= ? AND TM.MovementDateTime < ?
                               )";



                        using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                        {
                            wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            int count = (int)wardAdmissionsCmd.ExecuteScalar();

                                label8.Text = count.ToString();
                             }
                    }

                    //Deaths only
                    else if (category == "Death")
                    {
                        string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ? AND TM.ToWard = ? AND PM.dischargeDateTime >= ? AND PM.dischargeDateTime < ?
                               )";



                        using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                        {
                            wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = startDate });
                            wardAdmissionsCmd.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = exclusiveEndDate });

                            int count = (int)wardAdmissionsCmd.ExecuteScalar();

                                label9.Text = count.ToString();
                       
                    }
                   
                    }
                   
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error occurred\n" + error.Message,
                                    "Marondera Provincial Hospital",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
               }
            }
        

        private void button6_Click(object sender, EventArgs e)
        {
            label5.Text = "00";
            label6.Text = "00";
            label7.Text = "00";
            label8.Text = "00";
            label9.Text = "00";
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }
    }
}