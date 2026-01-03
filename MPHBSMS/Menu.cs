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
            const string CatTransferIn = "InterWardTransferIn";
            const string CatDischarge = "Discharge";
            const string CatTransferOut = "InterWardTransferOut";
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
                string name = textBox2.Text.Trim().ToUpper();
                string surname = textBox3.Text.Trim().ToUpper();
                string enteredBy = MPHBSMS.CurrentUser;
                DateTime movementDT = dateTimePicker1.Value;

                // Determine Gender
                string gender = "";
                if (radioButton1.Checked) gender = "FEMALE";
                else if (radioButton2.Checked) gender = "MALE";

                // 2. PRE-SAVE VALIDATION
                if (movementDT > DateTime.Now)
                {
                    MessageBox.Show("Future dates are not allowed.", "Invalid Date", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!System.Text.RegularExpressions.Regex.IsMatch(hospitalNumber, @"^\d+\/\d+$"))
                {
                    MessageBox.Show("Invalid format. Use 000/00.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(hospitalNumber) || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(category) || !checkBox6.Checked)
                {
                    MessageBox.Show("Please complete all required fields.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 3. EXISTENCE CHECK
                bool patientExists = false;
                string checkQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE [hospitalNumber] = ?";
                using (OleDbCommand checkCmd = new OleDbCommand(checkQuery, con))
                {
                    checkCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                    patientExists = Convert.ToInt32(checkCmd.ExecuteScalar()) > 0;
                }

                if (category == CatAdmission && patientExists)
                {
                    MessageBox.Show("Patient is already admitted.", "Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (category != CatAdmission && !patientExists)
                {
                    MessageBox.Show("Patient record not found.", "Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 4. START ATOMIC TRANSACTION
                transaction = con.BeginTransaction();
                string movementIDnum = MovementIdGenerator.GenerateMovementId();

                // A. LOG MOVEMENT (Standard for all actions)
                string insertMovementQuery = "INSERT INTO tblPatientMovement " +
                    "([hospitalNumber],[name],[surname],[gender],[ToWard],[FromWard],[MovementDateTime],[category],[enteredBy],[movementID]) " +
                    "VALUES (?,?,?,?,?,?,?,?,?,?)";

                using (OleDbCommand movCmd = new OleDbCommand(insertMovementQuery, con, transaction))
                {
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = name;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = surname;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = gender;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = ward;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = fromLocation;
                    movCmd.Parameters.Add("?", OleDbType.Date).Value = movementDT;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = category;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = enteredBy;
                    movCmd.Parameters.Add("?", OleDbType.VarChar).Value = movementIDnum;
                    movCmd.ExecuteNonQuery();
                }

                // B. UPDATE MASTER TABLE
                if (category == CatAdmission)
                {
                    string insertMasterQuery = "INSERT INTO tblPatientMaster " +
                        "([hospitalNumber],[name],[surname],[gender],[CurrentWard],[IsAdmitted],[AdmissionDate]) " +
                        "VALUES (?,?,?,?,?,?,?)";
                    using (OleDbCommand masterCmd = new OleDbCommand(insertMasterQuery, con, transaction))
                    {
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = name;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = surname;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = gender;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = ward;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = "Yes";
                        masterCmd.Parameters.Add("?", OleDbType.Date).Value = movementDT;
                        masterCmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string finalWard = ward;
                    string isAdmitted = "Yes";

                    if (category == CatDischarge) { finalWard = "DISCHARGED"; isAdmitted = "No"; }
                    else if (category == CatDeath) { finalWard = "DECEASED"; isAdmitted = "No"; }

                    string updateMasterQuery = "UPDATE tblPatientMaster SET CurrentWard = ?, IsAdmitted = ?, DischargeDate = ? WHERE [hospitalNumber] = ?";

                    using (OleDbCommand masterCmd = new OleDbCommand(updateMasterQuery, con, transaction))
                    {
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = finalWard;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = isAdmitted;
                        masterCmd.Parameters.Add("?", OleDbType.Date).Value = (isAdmitted == "No") ? (object)movementDT : DBNull.Value;
                        masterCmd.Parameters.Add("?", OleDbType.VarChar).Value = hospitalNumber;
                        masterCmd.ExecuteNonQuery();
                    }
                }

                // 5. FINALIZE
                transaction.Commit();

                // Log the activity using standard concatenation to avoid $ interpolation
                MPHBSMS.LogActivity("PROCESS (" + category + "): Hospital No: " + hospitalNumber + " by " + enteredBy);

                ClearInputFields();
                MessageBox.Show("Patient record successfully processed.", "Marondera Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception error)
            {
                if (transaction != null) transaction.Rollback();
                MessageBox.Show("Database Error: " + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
            using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery ,con))
            {
                const string category = "Admission";
                wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                label16.Text = admissionInWard.ToString();
            }
            
                 //     interward transfer In only
            string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
            using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                        {
                            const string category = "Inter Ward Transfer In";
                            wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                            wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                            int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                            label17.Text = interwardTransferInInWard.ToString();
                        }

                     // discharge only
            // Discharge Count Query using DISTINCT on hospitalNumber
            string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

            using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
            {
                const string category = "Discharge";
                wardDischargeCmd.Parameters.AddWithValue("?", category);
                wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                label18.Text = dischargeInWard.ToString();
            }
                       // interward Transfer Out only
                        string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
            using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                        {
                            const string category = "Inter Ward Transfer Out";
                            wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                            wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                            int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                            label19.Text = interwardTransferOutInWard.ToString();
                        }


                        // death only
            string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
            using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                        {
                            const string category = "Death";
                            wardDeathCmd.Parameters.AddWithValue("?", category);
                            wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                            int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                            label20.Text = deathInWard.ToString();
                        }

                        // all beds occupied
            // Assuming selectedWard is defined and checked for null
            // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
            string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

            using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
            {
                const string category = "Discharge";
                wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                label21.Text = dischargeInWard.ToString();
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
    
        }

        private void postNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    } 
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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


                    // all admissions
                    string wardOccupancyQuery = "SELECT COUNT(*) FROM tblPatientMaster WHERE IsAdmitted = TRUE AND CurrentWard = ?";
                    using (OleDbCommand wardOccupancyCmd = new OleDbCommand(wardOccupancyQuery, con))
                    {
                        wardOccupancyCmd.Parameters.AddWithValue("?", selectedWard);
                        int bedsOccupiedInWard = (int)wardOccupancyCmd.ExecuteScalar();
                        label15.Text = bedsOccupiedInWard.ToString();
                    }

                    // admission only
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
    
        }

        private void labourWardToolStripMenuItem_Click(object sender, EventArgs e)
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    string wardAdmissionsQuery = @"SELECT COUNT(HospitalID) AS TotalAdmissions 
                               FROM (
                                   SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                   FROM tblPatientMaster AS PM 
                                   INNER JOIN tblPatientMovement AS TM 
                                   ON PM.hospitalNumber = TM.hospitalNumber 
                                   WHERE TM.category = ? AND PM.currentWard = ?
                               )";
                    using (OleDbCommand wardAdmissionsCmd = new OleDbCommand(wardAdmissionsQuery, con))
                    {
                        const string category = "Admission";
                        wardAdmissionsCmd.Parameters.AddWithValue("?", category);
                        wardAdmissionsCmd.Parameters.AddWithValue("?", selectedWard);
                        int admissionInWard = (int)wardAdmissionsCmd.ExecuteScalar();
                        label16.Text = admissionInWard.ToString();
                    }

                    //     interward transfer In only
                    string wardInterwardTransferInQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersIn 
                                        FROM (
                                            SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                            FROM tblPatientMaster AS PM 
                                            INNER JOIN tblPatientMovement AS TM 
                                            ON PM.hospitalNumber = TM.hospitalNumber 
                                            WHERE TM.category = ? AND PM.currentWard = ?
                                        )";
                    using (OleDbCommand wardInterwardTransferInCmd = new OleDbCommand(wardInterwardTransferInQuery, con))
                    {
                        const string category = "Inter Ward Transfer In";
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferInCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferInInWard = (int)wardInterwardTransferInCmd.ExecuteScalar();
                        label17.Text = interwardTransferInInWard.ToString();
                    }

                    // discharge only
                    // Discharge Count Query using DISTINCT on hospitalNumber
                    string wardDischargeQuery = @"SELECT COUNT(HospitalID) AS DischargeCount
                              FROM (
                                  SELECT DISTINCT hospitalNumber AS HospitalID 
                                  FROM tblPatientMovement 
                                  WHERE category = ? AND FromWard = ?
                              )";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(wardDischargeQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", category);
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label18.Text = dischargeInWard.ToString();
                    }
                    // interward Transfer Out only
                    string wardInterwardTransferOutQuery = @"SELECT COUNT(HospitalID) AS TotalTransfersOut 
                                         FROM (
                                             SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                             FROM tblPatientMaster AS PM 
                                             INNER JOIN tblPatientMovement AS TM 
                                             ON PM.hospitalNumber = TM.hospitalNumber 
                                             WHERE TM.category = ? AND PM.currentWard = ?
                                         )";
                    using (OleDbCommand wardInterwardTransferOutCmd = new OleDbCommand(wardInterwardTransferOutQuery, con))
                    {
                        const string category = "Inter Ward Transfer Out";
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", category);
                        wardInterwardTransferOutCmd.Parameters.AddWithValue("?", selectedWard);
                        int interwardTransferOutInWard = (int)wardInterwardTransferOutCmd.ExecuteScalar();
                        label19.Text = interwardTransferOutInWard.ToString();
                    }


                    // death only
                    string wardDeathQuery = @"SELECT COUNT(HospitalID) AS TotalDeaths 
                          FROM (
                              SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                              FROM tblPatientMaster AS PM 
                              INNER JOIN tblPatientMovement AS TM 
                              ON PM.hospitalNumber = TM.hospitalNumber 
                              WHERE TM.category = ? AND PM.currentWard = ?
                          )";
                    using (OleDbCommand wardDeathCmd = new OleDbCommand(wardDeathQuery, con))
                    {
                        const string category = "Death";
                        wardDeathCmd.Parameters.AddWithValue("?", category);
                        wardDeathCmd.Parameters.AddWithValue("?", selectedWard);
                        int deathInWard = (int)wardDeathCmd.ExecuteScalar();
                        label20.Text = deathInWard.ToString();
                    }

                    // all beds occupied
                    // Assuming selectedWard is defined and checked for null
                    // FIX: Join PM and TM, then filter by PM.currentWard and TM.category using the subquery structure.
                    string combinedMovementQuery = @"SELECT COUNT(HospitalID) AS CombinedMovementCount
                                 FROM (
                                     SELECT DISTINCT TM.hospitalNumber AS HospitalID 
                                     FROM tblPatientMaster AS PM 
                                     INNER JOIN tblPatientMovement AS TM 
                                     ON PM.hospitalNumber = TM.hospitalNumber 
                                     WHERE PM.currentWard = ? 
                                     AND (TM.category = 'Admission' OR TM.category = 'InterWardTransferIn' OR TM.category = 'InterWardTransferOut')
                                 ) AS CombinedMovementResults";

                    using (OleDbCommand wardDischargeCmd = new OleDbCommand(combinedMovementQuery, con))
                    {
                        const string category = "Discharge";
                        wardDischargeCmd.Parameters.AddWithValue("?", selectedWard);
                        int dischargeInWard = (int)wardDischargeCmd.ExecuteScalar();
                        label21.Text = dischargeInWard.ToString();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
    }
}
