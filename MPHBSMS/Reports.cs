using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MPHBSMS
{
    public partial class Reports : Form
    {
        // 💡 INITIAL GREETING
      
          private ReportManager _manager;

        public Reports()
        {
            InitializeComponent();

            // 2. INITIALIZATION: Create a new instance of the manager
            _manager = new ReportManager();

            // Display the initial greeting and prompt
            DisplayBotGreeting();
            UpdateUI(initialCall: true);

        }

        private void Reports_Load(object sender, EventArgs e)
        {

        }

        private void laborWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Using textBox1 for user input
            string userInput = textBox1.Text.Trim();

            if (string.IsNullOrWhiteSpace(userInput))
            {
                // Using richTextBox1 for output
                ChatLogger.AppendText(richTextBox1, "Bot: Please enter a command.", Color.Gray);
                return;
            }

            // 1. Log the user's input
            ChatLogger.AppendText(richTextBox1, "You: " + userInput, Color.DarkBlue, true);

            // 2. Call the manager to process the input
            ProcessResult result = _manager.ProcessAndAdvance(userInput);

            if (!result.Success)
            {
                // If validation failed, log the error message
                ChatLogger.AppendText(richTextBox1, "MPH Bot Error: " + result.ErrorMessage, Color.Red);
            }
            else
            {
                // Log the successful state transition
                ChatLogger.AppendText(richTextBox1, "--- State Advanced to: " + _manager.CurrentState.ToString() + " ---", Color.OrangeRed, true);
            }

            // 3. Update UI based on the new state (This displays the next prompt or the final execution)
            UpdateUI();

            // 4. Clear input box
            textBox1.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Clear the existing log
            richTextBox1.Clear();

            // Re-instantiate the manager to reset all conversation parameters
            _manager = new ReportManager();

            // Restart the conversation flow
            DisplayBotGreeting();
            UpdateUI();
        }

        private void UpdateUI(bool initialCall = false)
        {
            // 1. Check for the final execution state
            if (_manager.CurrentState == ConversationState.ConfirmationAndGenerate)
            {
                string finalLog = _manager.ExecuteReportLogic();

                // Log the report execution result
                ChatLogger.AppendText(richTextBox1, "\n--- EXECUTION REPORT ---", Color.Gray, true);
                ChatLogger.AppendText(richTextBox1, finalLog, Color.DarkGreen, false);
                ChatLogger.AppendText(richTextBox1, "------------------------\n", Color.Gray, true);

                _manager.CurrentState = ConversationState.Complete;

                UpdateUI(); // Final call to display "Operation complete."
                return;
            }

            // 2. Display the bot's standard prompt for the new state
            string prompt = _manager.GetCurrentPrompt();

            // Display the prompt if we are past the initial greeting
            if (_manager.CurrentState != ConversationState.Greeting) 
            {
                ChatLogger.AppendText(richTextBox1, "Bot: " + prompt, Color.Green, false);
            }
        }

        private void DisplayBotGreeting()
        {
            string greeting = "Welcome to the Report Generator! I'm here to guide you through creating or opening a report.";
            ChatLogger.AppendText(richTextBox1, "Bot: " + greeting, Color.Green, true);
        }

        // 💡 5. HELPER METHODS FOR RICH TEXT BOX INPUT
        private string GetLastLineInput()
        {
            string fullText = richTextBox1.Text;
            int lastNewLine = fullText.LastIndexOf(Environment.NewLine);

            if (lastNewLine > 0)
            {
                // Return the text from the last new line to the end of the box
                return fullText.Substring(lastNewLine + Environment.NewLine.Length).Trim();
            }
            // Case for the very first command entered
            return fullText.Trim();
        }

        private void ClearLastLineInput()
        {
            // Deletes the user's input line after it has been read.
            int lastNewLine = richTextBox1.Text.LastIndexOf(Environment.NewLine);
            
            if (lastNewLine > 0)
            {
                // Select and delete the content from the start of the user's line to the end.
                richTextBox1.Select(lastNewLine + Environment.NewLine.Length, richTextBox1.Text.Length - (lastNewLine + Environment.NewLine.Length));
                richTextBox1.SelectedText = string.Empty;
            }
            // This handles the case where the user deletes the whole input but only enters the very first command.
            else
            {
                richTextBox1.Clear();
                // We must re-display the greeting after clearing the whole box
                DisplayBotGreeting();
            }
        }

       /* // === Placeholder Event Handlers (DELETE button4_Click as it's replaced) ===
        private void Reports_Load(object sender, EventArgs e) { }
        private void laborWardToolStripMenuItem_Click(object sender, EventArgs e) { }
        // The default richTextBox1_TextChanged is often empty and can be deleted if you have no other logic there.
        // private void richTextBox1_TextChanged(object sender, EventArgs e) { }*/
    

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

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
    }
}
