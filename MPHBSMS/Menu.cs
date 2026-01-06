using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MPHBSMS
{
    public partial class Menu : Form
    {
        public string selectedWard;
        private void ClearInputFields()
        {
            // Reset TextBoxes
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();

            // Reset Radio Buttons
            radioButton1.Checked = false;
            radioButton2.Checked = false;

            // Reset CheckBoxes
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;

            // Reset ListBoxes 
            if (listBox1.SelectedIndex != -1) listBox1.SelectedIndex = -1;
            if (listBox2.SelectedIndex != -1) listBox2.SelectedIndex = -1;

            // Reset DatePicker
            dateTimePicker1.Value = DateTime.Now;

            // Set focus back to the first input field
            textBox1.Focus();

            DateTime now = DateTime.Now;

            // Set your desired default
            dateTimePicker1.Value = now.AddDays(-1);
        }
        private string prevFromSelection = null;
        private string prevToSelection = null;

        private readonly HashSet<string> femaleOnly = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "Labor Ward",
    "Female Ward",
    "PostNatal Ward",
    "AnteNatal Ward"
};

        private const string maleOnly = "Male Ward";


        public Menu()
        {
            InitializeComponent();
            listBox1.Items.Add("Accident and Emergence");
            listBox1.Items.Add("High Dependency Unit");
            listBox1.Items.Add("Mental Health Unit");
            listBox1.Items.Add("Somewhwere Else");
            listBox1.Items.Add("Funeral Parlour");
            listBox1.Items.Add("AnteNatal Ward");
            listBox1.Items.Add("PostNatal Ward");
            listBox1.Items.Add("Paedatric Ward");
            listBox1.Items.Add("Other Hospital");
            listBox1.Items.Add("NeoNatal Ward");
            listBox1.Items.Add("Other Clinic");
            listBox1.Items.Add("Female Ward");
            listBox1.Items.Add("Labor Ward");
            listBox1.Items.Add("Male Ward");
            listBox1.Items.Add("Mortuary");
            listBox1.Items.Add("Home");
            listBox1.Items.Add("Null");

            listBox2.Items.Add("Accident and Emergence");
            listBox2.Items.Add("High Dependency Unit");
            listBox2.Items.Add("Mental Health Unit");
            listBox2.Items.Add("Somewhwere Else");
            listBox2.Items.Add("Funeral Parlour");
            listBox2.Items.Add("AnteNatal Ward");
            listBox2.Items.Add("PostNatal Ward");
            listBox2.Items.Add("Paedatric Ward");
            listBox2.Items.Add("Other Hospital");
            listBox2.Items.Add("NeoNatal Ward");
            listBox2.Items.Add("Other Clinic");
            listBox2.Items.Add("Female Ward");
            listBox2.Items.Add("Labor Ward");
            listBox2.Items.Add("Male Ward");
            listBox2.Items.Add("Mortuary");
            listBox2.Items.Add("Home");
            listBox2.Items.Add("Null");

            
           
        }

        private void Menu_Load(object sender, EventArgs e)
        {
            
                DatabaseHelper.InitializeDatabase();
                textBox1.Focus();

    DateTime now = DateTime.Now;

    // Set your desired default
    dateTimePicker1.Value = now.AddDays(-1);



                prevFromSelection = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString() : null;
                prevToSelection = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString() : null;

                // wire events if not already wired in Designer
                listBox2.SelectedIndexChanged -= listBox2_SelectedIndexChanged;
                listBox1.SelectedIndexChanged -= listBox1_SelectedIndexChanged;
                listBox2.SelectedIndexChanged += listBox2_SelectedIndexChanged;
                listBox1.SelectedIndexChanged += listBox1_SelectedIndexChanged;
            
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
            const string CatAdmission = "Admission";
            const string CatTransferIn = "Inter Ward Transfer In";
            const string CatDischarge = "Discharge";
            const string CatTransferOut = "Inter Ward Transfer Out";
            const string CatDeath = "Death";

            OleDbConnection con = null;
            OleDbTransaction transaction = null;

            try
            {
                con = new OleDbConnection(DatabaseHelper.ConnectionString);
                con.Open();

                // 1. DATA GATHERING
                string category = "";
                if (checkBox1.Checked) category = CatAdmission;
                else if (checkBox2.Checked) category = CatTransferIn;
                else if (checkBox3.Checked) category = CatDischarge;
                else if (checkBox4.Checked) category = CatTransferOut;
                else if (checkBox5.Checked) category = CatDeath;

                string ward = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString().Trim() : "";
                string fromLocation = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString().Trim() : "";
                string hospitalNumber = textBox1.Text.Trim().ToUpper();
                // Updated Code for Title Case
                var textInfo = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
                string name = textInfo.ToTitleCase(textBox2.Text.Trim().ToLower());
                string surname = textInfo.ToTitleCase(textBox3.Text.Trim().ToLower());
                string enteredBy = MPHBSMS.CurrentUser;
                DateTime movementDT = dateTimePicker1.Value;
                string gender = radioButton1.Checked ? "FEMALE" : (radioButton2.Checked ? "MALE" : "");

                transaction = con.BeginTransaction();

                // 2. STEP 1: SYNC THE MASTER STATUS (Must happen first for Relationship Integrity)
                string isAdmitted = (category == CatDischarge || category == CatDeath) ? "No" : "Yes";
                string currentStatus = ward;
                if (category == CatDischarge) currentStatus = "DISCHARGED";
                if (category == CatDeath) currentStatus = "DECEASED";

                bool masterExists = false;
                using (OleDbCommand cmd = new OleDbCommand("SELECT COUNT(*) FROM tblPatientMaster WHERE [hospitalNumber] = ?", con, transaction))
                {
                    cmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                    masterExists = Convert.ToInt32(cmd.ExecuteScalar()) > 0;
                }

                if (!masterExists)
                {
                    // Insert Master first so the Foreign Key relationship is satisfied
                    string insMaster = "INSERT INTO tblPatientMaster ([hospitalNumber],[name],[surname],[gender],[currentWard],[isAdmitted]) VALUES (?,?,?,?,?,?)";
                    using (OleDbCommand insCmd = new OleDbCommand(insMaster, con, transaction))
                    {
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = name;
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = surname;
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = gender;
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = currentStatus;
                        insCmd.Parameters.Add("?", OleDbType.VarChar).Value = isAdmitted;
                        insCmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string updMaster = "UPDATE tblPatientMaster SET [currentWard] = ?, [isAdmitted] = ? WHERE [hospitalNumber] = ?";
                    using (OleDbCommand updCmd = new OleDbCommand(updMaster, con, transaction))
                    {
                        updCmd.Parameters.Add("?", OleDbType.VarChar).Value = currentStatus;
                        updCmd.Parameters.Add("?", OleDbType.VarChar).Value = isAdmitted;
                        updCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                        updCmd.ExecuteNonQuery();
                    }
                }

                // 3. STEP 2: INSERT INTO HISTORY (tblPatientMovement)
                // FIX: Field names updated to match schema (MovementDateTime, toWard, fromWard)
                string movementIDnum = MovementIdGenerator.GenerateMovementId();
                string insertMovementQuery = "INSERT INTO tblPatientMovement " +
                    "([hospitalNumber],[name],[surname],[gender],[toWard],[fromWard],[MovementDateTime],[category],[enteredBy],[movementID]) " +
                    "VALUES (?,?,?,?,?,?,?,?,?,?)";

                using (OleDbCommand movCmd = new OleDbCommand(insertMovementQuery, con, transaction))
                {
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = name;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = surname;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = gender;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = ward;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = fromLocation;

                    // Use NULL for New Admission timestamp as requested, otherwise use picker value
                    if (category == CatAdmission)
                        movCmd.Parameters.Add("?", OleDbType.Date).Value = DBNull.Value;
                    else
                        movCmd.Parameters.Add("?", OleDbType.Date).Value = movementDT;

                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = category;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = enteredBy;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = movementIDnum;

                    movCmd.ExecuteNonQuery();
                }

                transaction.Commit();
                MessageBox.Show("Inserted successfully.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearInputFields();
            }
            catch (Exception error)
            {
                if (transaction != null) transaction.Rollback();
                MessageBox.Show("Process Failed: " + error.Message);
            }
            finally
            {
                if (con != null) con.Close();
            }
            }    


private void mentalHealthUnitToolStripMenuItem_Click(object sender, EventArgs e)
{

    ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
    selectedWard = ClickedItem.Text;

    // User Confirmation
    MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        try
        {
            con.Open();
            textBox1.Focus();

            // 1. Current Occupancy (Live status from Master table)
            string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
            using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
            {
                cmd.Parameters.AddWithValue("?", selectedWard);
                label15.Text = cmd.ExecuteScalar().ToString();
            }

            // 2. Admissions only
            string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                        FROM tblPatientMaster AS PM 
                                        INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                        WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Admission");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label16.Text = cmd.ExecuteScalar().ToString();
            }

            // 3. Interward Transfer In only
            string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                                SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                FROM tblPatientMaster AS PM 
                                                INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferInQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label17.Text = cmd.ExecuteScalar().ToString();
            }

            // 4. Discharge only (Checked via Movement table and FromWard)
            string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Discharge");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label18.Text = cmd.ExecuteScalar().ToString();
            }

            // 5. Interward Transfer Out only
            string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                                 SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                 FROM tblPatientMaster AS PM 
                                                 INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                 WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label19.Text = cmd.ExecuteScalar().ToString();
            }

            // 6. Death only
            string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                  FROM tblPatientMaster AS PM 
                                  INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                  WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Death");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label20.Text = cmd.ExecuteScalar().ToString();
            }

            // 7. Combined Movement (Total Traffic)
            // Note: For OLEDB, we use PM.currentWard = ? as the parameter.
            string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE PM.currentWard = ? 
                                             AND (TM.category = 'Admission' OR TM.category = 'Inter Ward Transfer In' OR TM.category = 'Inter Ward Transfer Out')
                                         ) AS CombinedMovementResults";

            using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
            {
                cmd.Parameters.AddWithValue("?", selectedWard);
                label21.Text = cmd.ExecuteScalar().ToString();
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

    // User Confirmation
    MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        try
        {
            con.Open();
            textBox1.Focus();

            // 1. Current Occupancy (Live status from Master table)
            string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
            using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
            {
                cmd.Parameters.AddWithValue("?", selectedWard);
                label15.Text = cmd.ExecuteScalar().ToString();
            }

            // 2. Admissions only
            string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                        FROM tblPatientMaster AS PM 
                                        INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                        WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Admission");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label16.Text = cmd.ExecuteScalar().ToString();
            }

            // 3. Interward Transfer In only
            string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                                SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                FROM tblPatientMaster AS PM 
                                                INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferInQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label17.Text = cmd.ExecuteScalar().ToString();
            }

            // 4. Discharge only (Checked via Movement table and FromWard)
            string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Discharge");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label18.Text = cmd.ExecuteScalar().ToString();
            }

            // 5. Interward Transfer Out only
            string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                                 SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                 FROM tblPatientMaster AS PM 
                                                 INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                 WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label19.Text = cmd.ExecuteScalar().ToString();
            }

            // 6. Death only
            string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                  FROM tblPatientMaster AS PM 
                                  INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                  WHERE TM.category = ? AND PM.currentWard = ?)";
            using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
            {
                cmd.Parameters.AddWithValue("?", "Death");
                cmd.Parameters.AddWithValue("?", selectedWard);
                label20.Text = cmd.ExecuteScalar().ToString();
            }

            // 7. Combined Movement (Total Traffic)
            // Note: For OLEDB, we use PM.currentWard = ? as the parameter.
            string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE PM.currentWard = ? 
                                             AND (TM.category = 'Admission' OR TM.category = 'Inter Ward Transfer In' OR TM.category = 'Inter Ward Transfer Out')
                                         ) AS CombinedMovementResults";

            using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
            {
                cmd.Parameters.AddWithValue("?", selectedWard);
                label21.Text = cmd.ExecuteScalar().ToString();
            }
        }
        catch (Exception error)
        {
            MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
        }

        private void paedatricWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;

            // Selection Confirmation
            MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Current Occupancy (Live Status)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                        FROM tblPatientMaster AS PM 
                                        INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                        WHERE TM.category = ? AND PM.currentWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In Only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                                SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                FROM tblPatientMaster AS PM 
                                                INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                WHERE TM.category = ? AND PM.currentWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Uses FromWard logic)
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out Only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                                 SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                                 FROM tblPatientMaster AS PM 
                                                 INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                                 WHERE TM.category = ? AND PM.currentWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Death Only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                  FROM tblPatientMaster AS PM 
                                  INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                  WHERE TM.category = ? AND PM.currentWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Movement Count
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE PM.currentWard = ? 
                                             AND (TM.category = 'Admission' OR TM.category = 'Inter Ward Transfer In' OR TM.category = 'Inter Ward Transfer Out')
                                         ) AS CombinedMovementResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } 
        }

        private void maleWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult result = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Current Occupancy (Live Census)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (Patients who entered the hospital through this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement AS TM 
                                        WHERE TM.category = ? AND TM.ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Patients moving into this ward from another)
                    string transferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                    SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                    FROM tblPatientMovement AS TM 
                                    WHERE TM.category = ? AND TM.ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Patients who left the hospital from this ward)
                    string dischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                   SELECT DISTINCT hospitalNumber AS HospitalID 
                                   FROM tblPatientMovement 
                                   WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(dischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Patients leaving this ward for another)
                    string transferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                     SELECT DISTINCT hospitalNumber AS HospitalID 
                                     FROM tblPatientMovement 
                                     WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Death (Recorded against the ward where it occurred)
                    string deathQuery = @"SELECT COUNT(HospitalID) FROM (
                               SELECT DISTINCT hospitalNumber AS HospitalID 
                               FROM tblPatientMovement 
                               WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(deathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Movement (Total Traffic / Activity)
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                          SELECT DISTINCT hospitalNumber AS HospitalID 
                                          FROM tblPatientMovement 
                                          WHERE (category IN ('Admission', 'Inter Ward Transfer In', 'Inter Ward Transfer Out', 'Discharge', 'Death')) 
                                          AND (FromWard = ? OR ToWard = ?)
                                        ) AS CombinedResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void postNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;

            // User Confirmation
            DialogResult result = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Live Occupancy (Who is physically in the ward right now)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (New arrivals to the hospital via this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Moved to this ward from another ward)
                    string wardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Patients who left the hospital from this ward)
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Moved from this ward to another ward)
                    string wardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                         SELECT DISTINCT hospitalNumber AS HospitalID 
                                         FROM tblPatientMovement 
                                         WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Deaths Only (Recorded from this ward)
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Movement (Total Activity: All Ins and Outs)
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                             SELECT DISTINCT hospitalNumber AS HospitalID 
                                             FROM tblPatientMovement 
                                             WHERE (FromWard = ? OR ToWard = ?)
                                         ) AS CombinedMovementResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        // Two '?' placeholders mean we add the parameter twice for OleDb
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void neoNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;

            // User Confirmation
            MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Live Occupancy (Who is physically in the ward right now)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (Entries into the hospital via this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Patients moving from another ward into this one)
                    string transferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                    SELECT DISTINCT hospitalNumber AS HospitalID 
                                    FROM tblPatientMovement 
                                    WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Patients leaving the hospital from this ward)
                    string dischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                   SELECT DISTINCT hospitalNumber AS HospitalID 
                                   FROM tblPatientMovement 
                                   WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(dischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Patients leaving this ward for another)
                    string transferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                     SELECT DISTINCT hospitalNumber AS HospitalID 
                                     FROM tblPatientMovement 
                                     WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Death Only
                    string deathQuery = @"SELECT COUNT(HospitalID) FROM (
                               SELECT DISTINCT hospitalNumber AS HospitalID 
                               FROM tblPatientMovement 
                               WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(deathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Total Traffic (Combined unique patients seen in this ward)
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                          SELECT DISTINCT hospitalNumber AS HospitalID 
                                          FROM tblPatientMovement 
                                          WHERE (FromWard = ? OR ToWard = ?)
                                        ) AS CombinedResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        // OleDb uses positional parameters: 1st ? is FromWard, 2nd ? is ToWard
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void anteNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Current Occupancy (Live Census)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (New hospital entries via this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Patients moving into this ward from another)
                    string wardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Patients leaving the hospital from this ward)
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Patients moving from this ward to another)
                    string wardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                         SELECT DISTINCT hospitalNumber AS HospitalID 
                                         FROM tblPatientMovement 
                                         WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Deaths (Recorded in this ward)
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Movement (Total Activity: Any movement associated with this ward)
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                         SELECT DISTINCT hospitalNumber AS HospitalID 
                                         FROM tblPatientMovement 
                                         WHERE FromWard = ? OR ToWard = ?
                                         ) AS CombinedMovementResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        // Two '?' markers mean we provide the ward twice
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void labourWardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;

            // User Confirmation
            MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Current Ward Occupancy (Live Count)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (New hospital arrivals to this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Moved from another ward into this one)
                    string wardTransferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Left hospital from this ward)
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                      SELECT DISTINCT hospitalNumber AS HospitalID 
                                      FROM tblPatientMovement 
                                      WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Moved from this ward to another)
                    string wardTransferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                         SELECT DISTINCT hospitalNumber AS HospitalID 
                                         FROM tblPatientMovement 
                                         WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardTransferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Death Only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Traffic (Total unique patients seen in this ward)
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) FROM (
                                         SELECT DISTINCT hospitalNumber AS HospitalID 
                                         FROM tblPatientMovement 
                                         WHERE (FromWard = ? OR ToWard = ?)
                                         ) AS CombinedResults";
                    using (OleDbCommand cmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        // First ? is for FromWard, Second ? is for ToWard
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving ward statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
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

            MPHBSMS.LogActivity("NAVIGATE: User opened the Visuals/Graphs section.");
            Visual obj = new Visual();
            this.Close();
            obj.Show();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            MPHBSMS.LogActivity("NAVIGATE: User returned to Home Dashboard.");

            // 1. Record the logout in the audit trail
            MPHBSMS.LogLogout();

            // 2. Clear the current user for security
            MPHBSMS.CurrentUser = "";

            Form1 obj = new Form1();
            this.Close();
            obj.Show();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Inline ward -> gender logic (null-safe)
            string ward = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString().Trim() : "";
            string gender = ""; // will be set below

            // Default: show both radios (will be adjusted for gender-specific wards)
            radioButton1.Visible = true;  // female
            radioButton2.Visible = true;  // male

            // Normalize for comparison
            string w = ward.ToLowerInvariant();

            if (w == "male ward")
            {
                // Male-only
                radioButton1.Visible = false;   // hide female
                radioButton2.Visible = true;
                radioButton2.Checked = true;
                gender = "male";
            }
            else if (w == "labor ward" ||
                     w == "female ward" ||
                     w == "postnatal ward" ||
                     w == "post natal ward" ||
                     w == "antenatal ward" ||
                     w == "ante natal ward")
            {
                // Female-only wards (handle common variants)
                radioButton2.Visible = false;   // hide male
                radioButton1.Visible = true;
                radioButton1.Checked = true;
                gender = "female";
            }
            else
            {
                // Mixed/unspecified ward: allow user to choose, preserve any existing checked state
                radioButton1.Visible = true;
                radioButton2.Visible = true;

                if (radioButton1.Checked) gender = "female";
                else if (radioButton2.Checked) gender = "male";
                else gender = ""; // force user to choose later during validation
            }
            string from = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString().Trim() : "";
    string to   = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString().Trim() : "";

    bool invalid = (string.Equals(from, maleOnly, StringComparison.OrdinalIgnoreCase) && femaleOnly.Contains(to))
                   || (femaleOnly.Contains(from) && string.Equals(to, maleOnly, StringComparison.OrdinalIgnoreCase));

    if (invalid)
    {
        MessageBox.Show("Cannot move patient from '" + from + "' to '" + to + "'.",
                        "Invalid Movement", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        // revert to previous valid selection if available
        if (!string.IsNullOrEmpty(prevToSelection) && listBox1.Items.Contains(prevToSelection))
            listBox1.SelectedItem = prevToSelection;
        else
            listBox1.SelectedIndex = -1;
    }
    else
    {
        // valid, remember it
        prevToSelection = to;
    }

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
            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // all admissions
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand wardOccupancyCmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        wardOccupancyCmd.Parameters.AddWithValue("?", selectedWard);
                        int bedsOccupiedInWard = (int)wardOccupancyCmd.ExecuteScalar();
                        label15.Text = bedsOccupiedInWard.ToString();
                    }

                    // admission only
                    string wardAdmissionsQuery = "SELECT COUNT(TM.hospitalNumber) AS TotalAdmissions FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE TM.category = ? AND PM.currentWard = ?";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = "SELECT COUNT(TM.hospitalNumber) AS TotalAdmissions FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE TM.category = ? AND PM.currentWard = ?";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    string wardDischargeQuery = "SELECT COUNT(TM.hospitalNumber) AS TotalAdmissions FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE TM.category = ? AND PM.currentWard = ?";
                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }

                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = "SELECT COUNT(TM.hospitalNumber) AS TotalAdmissions FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE TM.category = ? AND PM.currentWard = ?";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = "SELECT COUNT(TM.hospitalNumber) AS TotalAdmissions FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE TM.category = ? AND PM.currentWard = ?";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    string totalCensusQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand totalCensusCmd = new OleDbCommand(totalCensusQuery, con))
                    {
                        totalCensusCmd.Parameters.AddWithValue("?", selectedWard);
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

        private void accidentAndEmergenceToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;

            MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Live Occupancy (Patients physically in the ward right now)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions Only (New entries to the hospital via this ward)
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) FROM (
                                        SELECT DISTINCT hospitalNumber AS HospitalID 
                                        FROM tblPatientMovement 
                                        WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Admission");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (Patients moving from another ward into this one)
                    string transferInQuery = @"SELECT COUNT(HospitalID) FROM (
                                    SELECT DISTINCT hospitalNumber AS HospitalID 
                                    FROM tblPatientMovement 
                                    WHERE category = ? AND ToWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer In");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge Only (Patients leaving the hospital from this ward)
                    string dischargeQuery = @"SELECT COUNT(HospitalID) FROM (
                                   SELECT DISTINCT hospitalNumber AS HospitalID 
                                   FROM tblPatientMovement 
                                   WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(dischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Discharge");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Patients leaving this ward for another)
                    string transferOutQuery = @"SELECT COUNT(HospitalID) FROM (
                                     SELECT DISTINCT hospitalNumber AS HospitalID 
                                     FROM tblPatientMovement 
                                     WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(transferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Inter Ward Transfer Out");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Deaths Only
                    string deathQuery = @"SELECT COUNT(HospitalID) FROM (
                               SELECT DISTINCT hospitalNumber AS HospitalID 
                               FROM tblPatientMovement 
                               WHERE category = ? AND FromWard = ?)";
                    using (OleDbCommand cmd = new OleDbCommand(deathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", "Death");
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Traffic (Total unique patients processed in this ward)
                    string combinedQuery = @"SELECT COUNT(HospitalID) FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE FromWard = ? OR ToWard = ?
                                 ) AS TotalTraffic";
                    using (OleDbCommand cmd = new OleDbCommand(combinedQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }

                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Inline ward -> gender logic (null-safe)
            string ward = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString().Trim() : "";
            string gender = ""; // will be set below

            // Default: show both radios (will be adjusted for gender-specific wards)
            radioButton1.Visible = true;  // female
            radioButton2.Visible = true;  // male

            // Normalize for comparison
            string w = ward.ToLowerInvariant();

            if (w == "male ward")
            {
                // Male-only
                radioButton1.Visible = false;   // hide female
                radioButton2.Visible = true;
                radioButton2.Checked = true;
                gender = "male";
            }
            else if (w == "labor ward" ||
                     w == "female ward" ||
                     w == "postnatal ward" ||
                     w == "post natal ward" ||
                     w == "antenatal ward" ||
                     w == "ante natal ward")
            {
                // Female-only wards (handle common variants)
                radioButton2.Visible = false;   // hide male
                radioButton1.Visible = true;
                radioButton1.Checked = true;
                gender = "female";
            }
            else
            {
                // Mixed/unspecified ward: allow user to choose, preserve any existing checked state
                radioButton1.Visible = true;
                radioButton2.Visible = true;

                if (radioButton1.Checked) gender = "female";
                else if (radioButton2.Checked) gender = "male";
                else gender = ""; // force user to choose later during validation
            }
             string from = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString().Trim() : "";
    string to   = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString().Trim() : "";

    bool invalid = (string.Equals(from, maleOnly, StringComparison.OrdinalIgnoreCase) && femaleOnly.Contains(to))
                   || (femaleOnly.Contains(from) && string.Equals(to, maleOnly, StringComparison.OrdinalIgnoreCase));

    if (invalid)
    {
        MessageBox.Show("Cannot move patient from '" + from + "' to '" + to + "'.",
                        "Invalid Movement", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        // revert to previous valid selection if available
        if (!string.IsNullOrEmpty(prevFromSelection) && listBox2.Items.Contains(prevFromSelection))
            listBox2.SelectedItem = prevFromSelection;
        else
            listBox2.SelectedIndex = -1; // no valid previous, clear
    }
    else
    {
        // valid, remember it
        prevFromSelection = from;
    }

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Menu_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            MPHBSMS.LogActivity("NAVIGATE: User opened the Reports/Audit section.");
            Reports obj = new Reports();
            this.Close();
            obj.Show();
        }

        private void highDependencyUnitToolStripMenuItem_Click(object sender, EventArgs e)
        {

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    textBox1.Focus();

                    // 1. Current Occupancy (Live status)
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand cmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label15.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Admissions only
                    string wardAdmissionsQuery = @"SELECT COUNT(*) FROM (
                                        SELECT DISTINCT TM.hospitalNumber 
                                        FROM tblPatientMaster AS PM 
                                        INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                        WHERE TM.category = 'Admission' AND PM.currentWard = ?) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label16.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Interward Transfer In (FIXED THE JOIN ERROR HERE)
                    string transferInQuery = @"SELECT COUNT(*) FROM (
                                    SELECT DISTINCT TM.hospitalNumber 
                                    FROM tblPatientMaster AS PM 
                                    INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                                    WHERE TM.category = 'Inter Ward Transfer In' AND PM.currentWard = ?) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(transferInQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label17.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Discharge (Check FromWard)
                    string dischargeQuery = @"SELECT COUNT(*) FROM (
                                   SELECT DISTINCT hospitalNumber 
                                   FROM tblPatientMovement 
                                   WHERE category = 'Discharge' AND FromWard = ?) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(dischargeQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label18.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Interward Transfer Out (Check FromWard)
                    string transferOutQuery = @"SELECT COUNT(*) FROM (
                                     SELECT DISTINCT hospitalNumber 
                                     FROM tblPatientMovement 
                                     WHERE category = 'Inter Ward Transfer Out' AND FromWard = ?) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(transferOutQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label19.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 6. Death
                    string deathQuery = @"SELECT COUNT(*) FROM (
                               SELECT DISTINCT TM.hospitalNumber 
                               FROM tblPatientMaster AS PM 
                               INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber 
                               WHERE TM.category = 'Death' AND PM.currentWard = ?) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(deathQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label20.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 7. Combined Movement (Total Traffic)
                    string combinedQuery = @"SELECT COUNT(*) FROM (
                                  SELECT DISTINCT TM.hospitalNumber 
                                  FROM tblPatientMovement AS TM 
                                  WHERE (TM.category IN ('Admission', 'Inter Ward Transfer In', 'Inter Ward Transfer Out')) 
                                  AND (TM.FromWard = ? OR TM.ToWard = ?)
                                ) AS StatTable";
                    using (OleDbCommand cmd = new OleDbCommand(combinedQuery, con))
                    {
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        cmd.Parameters.AddWithValue("?", selectedWard);
                        label21.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void highDependencyUnitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
        }

        private void button6_Click(object sender, EventArgs e)
        {
   
    // 1. CONSTANTS FOR CATEGORIES
    const string CatTransferIn = "Inter Ward Transfer In";
    const string CatDischarge = "Discharge";
    const string CatTransferOut = "Inter Ward Transfer Out";
    const string CatDeath = "Death";

    OleDbConnection con = null;
    OleDbTransaction transaction = null;

    try
    {
        con = new OleDbConnection(DatabaseHelper.ConnectionString);
        con.Open();

        // 2. DATA GATHERING
        string hospitalNumber = textBox1.Text.Trim().ToUpper();
        
        // These are fetched as Read-Only; they are NOT updated in tblPatientMaster
        string name = textBox2.Text.Trim().ToUpper();
        string surname = textBox3.Text.Trim().ToUpper();
        string gender = radioButton1.Checked ? "FEMALE" : "MALE";
        
        string enteredBy = MPHBSMS.CurrentUser;
        DateTime movementDT = dateTimePicker1.Value;

        // Determine Movement Category
        string category = "";
        if (checkBox2.Checked) category = CatTransferIn;
        else if (checkBox3.Checked) category = CatDischarge;
        else if (checkBox4.Checked) category = CatTransferOut;
        else if (checkBox5.Checked) category = CatDeath;

        string toWard = listBox1.SelectedItem != null ? listBox1.SelectedItem.ToString().Trim() : "";
        string fromWard = listBox2.SelectedItem != null ? listBox2.SelectedItem.ToString().Trim() : "";

        // Validation
        if (string.IsNullOrEmpty(hospitalNumber) || string.IsNullOrEmpty(category))
        {
            MessageBox.Show("Please ensure a patient is loaded and a category is selected.", "Marondera Provincial Hospital");
            return;
        }

        // 3. START ATOMIC TRANSACTION
        transaction = con.BeginTransaction();
        string movementIDnum = MovementIdGenerator.GenerateMovementId();

        // Status Logic
        string isAdmittedStatus = (category == CatDischarge || category == CatDeath) ? "No" : "Yes";
        string currentWardPointer = toWard;
        if (category == CatDischarge) currentWardPointer = "DISCHARGED";
        if (category == CatDeath) currentWardPointer = "DECEASED";

        // 4. ACTION A: UPDATE MASTER (Only currentWard and isAdmitted)
        // Name, Surname, and Gender are NOT changed here to protect identity integrity.
        string updMaster = "UPDATE tblPatientMaster SET [currentWard] = ?, [isAdmitted] = ? WHERE [hospitalNumber] = ?";
        using (OleDbCommand updCmd = new OleDbCommand(updMaster, con, transaction))
        {
            updCmd.Parameters.Add("?", OleDbType.VarChar).Value = currentWardPointer;
            updCmd.Parameters.Add("?", OleDbType.VarChar).Value = isAdmittedStatus;
            updCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
            updCmd.ExecuteNonQuery();
        }

        // 5. ACTION B: INSERT NEW HISTORY RECORD (Audit Trace)
        string insertMovementQuery = "INSERT INTO tblPatientMovement " +
            "([hospitalNumber],[name],[surname],[gender],[toWard],[fromWard],[MovementDateTime],[category],[enteredBy],[movementID]) " +
            "VALUES (?,?,?,?,?,?,?,?,?,?)";

        using (OleDbCommand movCmd = new OleDbCommand(insertMovementQuery, con, transaction))
        {
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = name;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = surname;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = gender;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = toWard;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = fromWard;
            movCmd.Parameters.Add("?", OleDbType.Date).Value = movementDT;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = category;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = enteredBy;
            movCmd.Parameters.Add("?", OleDbType.VarChar).Value = movementIDnum;
            movCmd.ExecuteNonQuery();
        }

        transaction.Commit();
        MessageBox.Show("Patient movement recorded successfully.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

        // 6. CLEAR FIELDS (Incorporated into Main Code)
        textBox1.Clear(); 
        textBox2.Clear(); 
        textBox3.Clear();
        textBox2.ReadOnly = false;
        textBox3.ReadOnly = false;

        radioButton1.Checked = false;
        radioButton2.Checked = false;
        radioButton1.Enabled = true;
        radioButton2.Enabled = true;

        listBox1.SelectedIndex = -1;
        listBox2.SelectedIndex = -1;

        checkBox1.Checked = false;
        checkBox2.Checked = false;
        checkBox3.Checked = false;
        checkBox4.Checked = false;
        checkBox5.Checked = false;
        
        checkBox1.Enabled = true; 
        checkBox2.Enabled = false;
        checkBox3.Enabled = false;
        checkBox4.Enabled = false;
        checkBox5.Enabled = false;

        dateTimePicker1.Value = DateTime.Now;
        button6.Enabled = false; // Lock button until next fetch
        textBox1.Focus();

    }
    catch (Exception error)
    {
        if (transaction != null) transaction.Rollback();
        MessageBox.Show("Process Failed: " + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
    finally
    {
        if (con != null) con.Close();
    }
}
        

        private void button5_Click(object sender, EventArgs e)
        {
            string hospitalNumber = textBox1.Text.Trim().ToUpper();

            if (string.IsNullOrEmpty(hospitalNumber))
            {
                MessageBox.Show("Please enter a Hospital Number.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    string query = "SELECT * FROM tblPatientMaster WHERE hospitalNumber = ?";
                    using (OleDbCommand cmd = new OleDbCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("?", hospitalNumber);
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // 1. EXTRACT DATA
                                string currentWard = reader["currentWard"] != DBNull.Value ? reader["currentWard"].ToString().Trim() : "";
                                string isAdmitted = reader["isAdmitted"] != DBNull.Value ? reader["isAdmitted"].ToString() : "";

                                // 2. ADMISSION DATE: Pull from database to UI
                                if (reader["AdmissionDate"] != DBNull.Value)
                                {
                                    dateTimePicker1.Value = Convert.ToDateTime(reader["AdmissionDate"]);
                                }

                                // 3. STATUS VALIDATION
                                if (isAdmitted == "No" || currentWard == "DISCHARGED" || currentWard == "DECEASED")
                                {
                                    MessageBox.Show("Patient Status: " + currentWard + ". File is closed.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    return;
                                }

                                // 4. LOCK IDENTITY (No changes allowed to demographics)
                                textBox2.Text = reader["name"].ToString();
                                textBox3.Text = reader["surname"].ToString();
                                textBox2.ReadOnly = true;
                                textBox3.ReadOnly = true;

                                string gender = reader["gender"].ToString();
                                radioButton1.Checked = (gender == "FEMALE");
                                radioButton2.Checked = (gender == "MALE");
                                radioButton1.Enabled = false;
                                radioButton2.Enabled = false;

                                // 5. WARD HIGHLIGHTING
                                int wardIndexFrom = listBox2.FindStringExact(currentWard);
                                if (wardIndexFrom != -1) listBox2.SelectedIndex = wardIndexFrom;

                                int wardIndexTo = listBox1.FindStringExact(currentWard);
                                if (wardIndexTo != -1) listBox1.SelectedIndex = wardIndexTo;

                                // 6. CATEGORY ENFORCEMENT
                                checkBox1.Enabled = false; // Cannot re-admit an active patient
                                checkBox1.Checked = false;

                                // Enable categories for existing patients
                                checkBox2.Enabled = true; // Transfer In
                                checkBox3.Enabled = true; // Discharge
                                checkBox4.Enabled = true; // Transfer Out
                                checkBox5.Enabled = true; // Death

                                listBox2.Enabled = false; // "From" is fixed from DB
                                listBox1.Enabled = true;  // "To" is open for selection

                                button6.Enabled = true;

                                MessageBox.Show("Record Found. Select a movement category to proceed.", "Marondera Provincial Hospital");
                            }
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (!string.IsNullOrWhiteSpace(tb.Text))
            {
                // Converts "JOHN DOE" or "john doe" to "John Doe"
                tb.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(tb.Text.ToLower());
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (!string.IsNullOrWhiteSpace(tb.Text))
            {
                // Converts "JOHN DOE" or "john doe" to "John Doe"
                tb.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(tb.Text.ToLower());
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Allow letters, whitespace, and backspace
            if (!char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                // Stop the character from being entered into the control
                e.Handled = true;

                // Optional: Alert the user
                System.Media.SystemSounds.Beep.Play();
            }
        }
    }
}
