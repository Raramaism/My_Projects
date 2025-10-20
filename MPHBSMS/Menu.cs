using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MPHBSMS
{
    public partial class Menu : Form
    {
        public string selectedWard;


        public Menu()
        {
            InitializeComponent();
            listBox1.Items.Add("Mental Health Unit");
            listBox1.Items.Add("Female Ward");
            listBox1.Items.Add("Paedatric Ward");
            listBox1.Items.Add("Male Ward");
            listBox1.Items.Add("PostNatal Ward");
            listBox1.Items.Add("NeoNatal Ward");
            listBox1.Items.Add("AnteNatal Ward");
            listBox1.Items.Add("Labor Ward");
            listBox1.Items.Add("Accident and Emergence");
            listBox1.Items.Add("Null");

            listBox2.Items.Add("Home");
            listBox2.Items.Add("Other Hospital");
            listBox2.Items.Add("Other Clinics");
            listBox2.Items.Add("Somewhwere Else");
            listBox2.Items.Add("Mental Health Unit");
            listBox2.Items.Add("Female Ward");
            listBox2.Items.Add("Paedatric Ward");
            listBox2.Items.Add("Male Ward");
            listBox2.Items.Add("PostNatal Ward");
            listBox2.Items.Add("NeoNatal Ward");
            listBox2.Items.Add("AnteNatal Ward");
            listBox2.Items.Add("Labor Ward");
            listBox2.Items.Add("Accident and Emergence");
            listBox1.Items.Add("Null");
           
        }

        private void Menu_Load(object sender, EventArgs e)
        {
            try
            {
                DatabaseHelper.InitializeDatabase();
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    textBox1.Focus();

                    string query1 = ("Select Count(*)From mphtable");
                    OleDbCommand com1 = new OleDbCommand(query1, con);
                    int broughtForward = (int)com1.ExecuteScalar();
                    label15.Text = broughtForward.ToString();

                    string query7 = ("Select Count(*)From mphtable");
                    OleDbCommand com7 = new OleDbCommand(query7, con);
                    int bedsOccupied = (int)com7.ExecuteScalar();
                    label21.Text = bedsOccupied.ToString();
                    con.Close();
                }
                
            }
            catch (Exception error)
            {

                DialogResult = MessageBox.Show("ERROR!!\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
           
            
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
           
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

    // Constants for Movement Categories (Already defined in your original code)
    const string CatAdmission = "Admission";
    const string CatTransferIn = "Inter Ward Transfer In";
    const string CatDischarge = "Discharge";
    const string CatTransferOut = "Inter Ward Transfer Out";
    const string CatDeath = "Death";

    OleDbConnection con = null;
    try
    {
        con = new OleDbConnection(DatabaseHelper.ConnectionString);
        con.Open();

        // 1. DATA GATHERING & VALIDATION (Re-using your existing validation logic)
        string gender = "";
        if (radioButton1.Checked) gender = "female";
        else if (radioButton2.Checked) gender = "male";

        string category = "";
        if (checkBox1.Checked) category = CatAdmission;
        else if (checkBox2.Checked) category = CatTransferIn;
        else if (checkBox3.Checked) category = CatDischarge;
        else if (checkBox4.Checked) category = CatTransferOut;
        else if (checkBox5.Checked) category = CatDeath;

        string ward = listBox1.SelectedItem.ToString();
        // Rename 'home' to 'fromLocation' to be clearer for transfers
        string fromLocation = listBox2.SelectedItem.ToString(); 

        DateTime movementDT = dateTimePicker1.Value; // Use the DateTime object directly
        string hospitalNumber = textBox1.Text.Trim().ToLower();
        string name = textBox2.Text.Trim().ToLower();
        string surname = textBox3.Text.Trim().ToLower();
        string enteredBy = MPHBSMS.CurrentUser;

        // Validation Checks (Keeping your original checks)
        if (string.IsNullOrWhiteSpace(hospitalNumber) || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(surname) ||
            string.IsNullOrEmpty(gender) || string.IsNullOrEmpty(category) || !checkBox6.Checked || string.IsNullOrEmpty(ward) || !checkBox5.Checked)
        {
            MessageBox.Show("Please complete all required fields and actions.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 2. PATIENT EXISTENCE CHECK (Querying the dedicated LIVE status table)
        bool patientExists = false;
        // NOTE: We now check against the new master table for LIVE status
        string checkQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE [hospitalNumber] = ?";
        using (OleDbCommand checkCmd = new OleDbCommand(checkQuery, con))
        {
            checkCmd.Parameters.AddWithValue("?", hospitalNumber);
            patientExists = (int)checkCmd.ExecuteScalar() > 0;
        }

        // Admission / Discharge Conflict Checks
        if (category == CatAdmission && patientExists)
        {
            MessageBox.Show("This patient is already admitted. Cannot re-admit.", "Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (category != CatAdmission && !patientExists)
        {
            MessageBox.Show("Patient does not exist in the active patient database. Cannot process Transfer/Discharge.", "Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // --- START TRANSACTION ---
        
        // 3. LOG THE MOVEMENT (Insert into the historical tracing table)
        // NOTE: Renamed mphtable to tblPatientMovement for clarity/structure
        string insertMovementQuery = "INSERT INTO tblPatientMovement ([hospitalNumber],[name],[surname],[gender],[ToWard],[FromWard],[MovementDateTime],[category],[enteredBy],[movementID]) " +
                                     "VALUES (?,?,?,?,?,?,?,?,?)";
        
        string movementID = textBox1.Text + "ID";
        using (OleDbCommand movementCmd = new OleDbCommand(insertMovementQuery, con))
        {
            movementCmd.Parameters.AddWithValue("?", hospitalNumber);
            movementCmd.Parameters.AddWithValue("?", name);
            movementCmd.Parameters.AddWithValue("?", surname);
            movementCmd.Parameters.AddWithValue("?", gender);
            movementCmd.Parameters.AddWithValue("?", ward);             // 'To' location
            movementCmd.Parameters.AddWithValue("?", fromLocation);    // 'From' location
            movementCmd.Parameters.AddWithValue("?", movementDT);      // Combined Date/Time
            movementCmd.Parameters.AddWithValue("?", category);
            movementCmd.Parameters.AddWithValue("?", enteredBy);
            movementCmd.Parameters.AddWithValue("?", movementID);
            
            movementCmd.ExecuteNonQuery();
        }
        
        // 4. UPDATE MASTER RECORD (Update the LIVE status for Bed Statistics/Tracing)
        
        string updateMasterQuery = string.Empty;

        if (category == CatAdmission)
        {
            // Admission: INSERT into the Master table
            updateMasterQuery = "INSERT INTO tblPatientMaster ([hospitalNumber],[name],[surname],[gender],[CurrentWard],[IsAdmitted],[AdmissionDate]) " +
                                "VALUES (?,?,?,?,?,?,?)";
            using (OleDbCommand masterCmd = new OleDbCommand(updateMasterQuery, con))
            {
                masterCmd.Parameters.AddWithValue("?", hospitalNumber);
                masterCmd.Parameters.AddWithValue("?", name);
                masterCmd.Parameters.AddWithValue("?", surname);
                masterCmd.Parameters.AddWithValue("?", gender);
                masterCmd.Parameters.AddWithValue("?", ward);
                masterCmd.Parameters.AddWithValue("?", true); // IsAdmitted = TRUE
                masterCmd.Parameters.AddWithValue("?", movementDT);
                masterCmd.ExecuteNonQuery();
            }
        }
        else if (category == CatTransferIn)
        {
            // Transfer In: UPDATE Master table with new CurrentWard
            updateMasterQuery = "UPDATE tblPatientMaster SET CurrentWard = ?, IsAdmitted = TRUE WHERE [hospitalNumber] = ?";
            using (OleDbCommand masterCmd = new OleDbCommand(updateMasterQuery, con))
            {
                masterCmd.Parameters.AddWithValue("?", ward);
                masterCmd.Parameters.AddWithValue("?", hospitalNumber);
                masterCmd.ExecuteNonQuery();
            }
        }
        else if (category == CatDischarge || category == CatDeath || category == CatTransferOut)
        {
            // Disposal Categories: UPDATE Master table to set status as NOT admitted
            updateMasterQuery = "UPDATE tblPatientMaster SET CurrentWard = NULL, IsAdmitted = FALSE, DischargeDate = ? WHERE [hospitalNumber] = ?";
            using (OleDbCommand masterCmd = new OleDbCommand(updateMasterQuery, con))
            {
                masterCmd.Parameters.AddWithValue("?", movementDT); // Record Discharge Date/Time
                masterCmd.Parameters.AddWithValue("?", hospitalNumber);
                masterCmd.ExecuteNonQuery();
            }
        }

        // --- END TRANSACTION ---

        MessageBox.Show("Record processed successfully by " + enteredBy,
                        "Marondera Provincial Hospital",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
    }
    catch (Exception error)
    {
        MessageBox.Show("ERROR!!\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
    }
    finally
    {
        if (con != null && con.State == System.Data.ConnectionState.Open)
        {
            con.Close();
        }
    }


        }    


private void mentalHealthUnitToolStripMenuItem_Click(object sender, EventArgs e)
{
    ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender;
    selectedWard = clickedItem.Text; // Set the selected ward
    
    // Display selected ward
    MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

    // DatabaseHelper.InitializeDatabase(); // Ensure this is handled at application start
    
    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        try
        {
            con.Open();
            textBox1.Focus();

            // 1. Get Occupancy for the SELECTED Ward (Corresponds to B/F and Box 5 logic)
            // We must query the LIVE tblPatientMaster table, not the movement log!
            string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
            using (OleDbCommand wardOccupancyCmd = new OleDbCommand(wardOccupancyQuery, con))
            {
                wardOccupancyCmd.Parameters.AddWithValue("?", selectedWard);
                int bedsOccupiedInWard = (int)wardOccupancyCmd.ExecuteScalar();
                
                // This value represents the current live census for this ward (similar to the paper form's Box 5 C/F)
                label15.Text = bedsOccupiedInWard.ToString(); 
            }

            // 2. Get Total Hospital Census (All Beds Occupied)
            // This is the total number of patients currently admitted in the entire hospital.
            string totalCensusQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE";
            using (OleDbCommand totalCensusCmd = new OleDbCommand(totalCensusQuery, con))
            {
                int bedsOccupiedTotal = (int)totalCensusCmd.ExecuteScalar();
                label21.Text = bedsOccupiedTotal.ToString(); 
            }
        }
        catch (Exception error)
        {
            MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

        }

        private void femaleWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void paedatricWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void maleWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void postNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void neoNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void anteNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void labourWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox currentCheckBox = (CheckBox)sender;

            if (currentCheckBox.Checked)
            {
                CheckBox[] checkBoxes = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };

                foreach (CheckBox cb in checkBoxes)
                {
                    if (cb != currentCheckBox)
                    {
                        cb.Checked = false;
                    }
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox currentCheckBox = (CheckBox)sender;

            if (currentCheckBox.Checked)
            {
                CheckBox[] checkBoxes = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };

                foreach (CheckBox cb in checkBoxes)
                {
                    if (cb != currentCheckBox)
                    {
                        cb.Checked = false;
                    }
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox currentCheckBox = (CheckBox)sender;

            if (currentCheckBox.Checked)
            {
                CheckBox[] checkBoxes = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };

                foreach (CheckBox cb in checkBoxes)
                {
                    if (cb != currentCheckBox)
                    {
                        cb.Checked = false;
                    }
                }
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox currentCheckBox = (CheckBox)sender;

            if (currentCheckBox.Checked)
            {
                CheckBox[] checkBoxes = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };

                foreach (CheckBox cb in checkBoxes)
                {
                    if (cb != currentCheckBox)
                    {
                        cb.Checked = false;
                    }
                }
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox currentCheckBox = (CheckBox)sender;

            if (currentCheckBox.Checked)
            {
                CheckBox[] checkBoxes = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };

                foreach (CheckBox cb in checkBoxes)
                {
                    if (cb != currentCheckBox)
                    {
                        cb.Checked = false;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Visual obj = new Visual();
            this.Close();
            obj.Show();

        
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 obj = new Form1();
            this.Close();
            obj.Show();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void accidentAndEmergenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void accidentAndEmergenceToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Menu_Click(object sender, EventArgs e)
        {

        }
    }
}
