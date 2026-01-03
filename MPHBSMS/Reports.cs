using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Microsoft.Reporting.WinForms;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MPHBSMS
{
    public partial class Reports : Form
    {
        
        public Reports()
        {
            InitializeComponent();

        }


        private void Reports_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("Admission");
            comboBox1.Items.Add("InterWardTransferIn");
            comboBox1.Items.Add("Discharge");
            comboBox1.Items.Add("InterWardTransferOut");
            comboBox1.Items.Add("Death");

            comboBox2.Items.Add("Accident and Emergence");
            comboBox2.Items.Add("High Dependency Unit");
            comboBox3.Items.Add("Mental Health Unit");
            comboBox2.Items.Add("AnteNatal Ward");
            comboBox2.Items.Add("PostNatal Ward");
            comboBox2.Items.Add("Paedatric Ward");
            comboBox2.Items.Add("Funeral Parlour");
            comboBox2.Items.Add("Female Ward");
            comboBox2.Items.Add("Labor Ward");
            comboBox2.Items.Add("Male Ward");
            comboBox2.Items.Add("Mortuary");

           UpdateBackupStatusLabel();

            // Define the hospital folder path
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(documentsPath, "Marondera Provincial Hospital Bed Statistics Reports");

            // Create the folder on startup if it doesn't exist
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Example using a simple array for binding
            string[] columns = { "Hospital Number", "Name", "Surname", "Category", "Ward", "Admission Date", "Discharge Date" };
            comboBox3.DataSource = columns;
            comboBox3.SelectedIndex = 0;


            // Use the full name to avoid iTextSharp conflict
            System.Collections.Generic.List<string> myLogs = MPHBSMS.GetAvailableLogFiles();

            listBox1.Items.Clear(); // Changed from listBoxLogs to listBox1

            foreach (string fileName in myLogs) // 'myLogs' must match the line above
            {
                listBox1.Items.Add(fileName);
            }



        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 1. Validation for search input only
            if (string.IsNullOrWhiteSpace(textBox16.Text))
            {
                MessageBox.Show("Please enter a value to search.", "Marondera Provincial Hospital");
                return;
            }

            try
            {
                string category = comboBox3.Text.Replace(" ", "");

                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // SQL with INNER JOIN to fill all textboxes from both tables
                    string sql = "SELECT PM.hospitalNumber, PM.name, PM.surname, PM.gender, PM.currentWard, PM.isAdmitted, " +
                                 "PM.admissionDate, PM.dischargeDate, TM.MovementDateTime, TM.toWard, TM.fromWard, " +
                                 "TM.category, TM.enteredBy, TM.movementID " +
                                 "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                 "WHERE PM." + category + " = ? " +
                                 "ORDER BY PM.hospitalNumber ASC";

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.Parameters.AddWithValue("@v", textBox16.Text);
                    OleDbDataReader r = cmd.ExecuteReader();

                    if (r.Read())
                    {
                        // Filling all textboxes in ascending order
                        textBox3.Text = r["hospitalNumber"].ToString();
                        textBox4.Text = r["name"].ToString();
                        textBox5.Text = r["surname"].ToString();
                        textBox6.Text = r["toWard"].ToString();
                        textBox7.Text = r["fromWard"].ToString();
                        textBox8.Text = r["gender"].ToString();
                        textBox9.Text = r["category"].ToString();
                        textBox10.Text = r["dischargeDate"].ToString();
                        textBox11.Text = r["isAdmitted"].ToString();
                        textBox12.Text = r["currentWard"].ToString();
                        textBox13.Text = r["admissionDate"].ToString();
                        textBox14.Text = r["enteredBy"].ToString();
                        textBox15.Text = r["MovementDateTime"].ToString();
                        textBox21.Text = r["movementID"].ToString();

                        // --- ROBUST BED DAYS CALCULATION ---
                        // We parse the exact format from your table: DD/MM/YYYY HH:mm:ss
                        DateTime admDate;
                        DateTime disDate;

                        // TryParse handles cases where dischargeDate might be null/empty in the table
                        bool hasAdm = DateTime.TryParse(textBox13.Text, out admDate);
                        bool hasDis = DateTime.TryParse(textBox10.Text, out disDate);

                        if (hasAdm && hasDis)
                        {
                            // Calculate difference based on Date only to ignore time-of-day bias
                            TimeSpan ts = disDate.Date - admDate.Date;
                            int days = ts.Days;

                            // Standard hospital logic: stay is at least 1 day
                            if (days <= 0) days = 1;

                            label17.Text = days.ToString("D2"); // Formats as '01', '05', etc.
                        }
                        else
                        {
                            // If patient hasn't been discharged yet, Bed Days remains 00 or current duration
                            label17.Text = "00";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No record found matching that criteria.", "Marondera Provincial Hospital");
                    }
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                // Custom error message format for Marondera Provincial Hospital
                MessageBox.Show("ERROR!!\n" + ex.Message,
                                "Marondera Provincial Hospital",
                                MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }


        private void button5_Click(object sender, EventArgs e)
        {

        }


        // === Navigation/Placeholder Event Handlers ===
        private void button1_Click(object sender, EventArgs e)
        {
            // 1. Record the logout in the audit trail
            MPHBSMS.LogLogout();

            // 2. Clear the current user for security
            MPHBSMS.CurrentUser = "";

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

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click_1(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            try
            {
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                textBox2.Clear();
                textBox1.Clear();
                label5.Text = "00";
                label6.Text = "00";
                label7.Text = "00";
                label8.Text = "00";
                label9.Text = "00";
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click_2(object sender, EventArgs e)
        {

            DateTime startDate, endDate;

            // Validate Dates
            if (!DateTime.TryParse(textBox1.Text, out startDate) ||
                !DateTime.TryParse(textBox2.Text, out endDate))
            {
                button1.Enabled = false;
                return;
            }

            button1.Enabled = true;

            // Ensure both Ward and Category are selected
            if (comboBox2.SelectedItem == null || comboBox1.SelectedItem == null)
                return;

            string ward = comboBox2.SelectedItem.ToString();
            string selectedCategory = comboBox1.SelectedItem.ToString();

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();
                    // Use switch to only update the category selected in comboBox1

                    switch (selectedCategory)
                    {
                        case "Admission":
                            using (OleDbCommand cmd = new OleDbCommand(
                                "SELECT COUNT(*) FROM tblPatientMaster WHERE currentWard = ? AND isAdmitted = 'Yes'", con))
                            {
                                cmd.Parameters.AddWithValue("?", ward);
                                label5.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                            }
                            break;

                        case "Discharge":
                            using (OleDbCommand cmd = new OleDbCommand(
                                "SELECT COUNT(*) FROM tblPatientMaster WHERE currentWard = ? AND isAdmitted = 'No'", con))
                            {
                                cmd.Parameters.AddWithValue("?", ward);
                                label7.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                            }
                            break;

                        case "InterWardTransferIn":
                            using (OleDbCommand cmd = new OleDbCommand(
                                @"SELECT COUNT(*) FROM (
                SELECT DISTINCT hospitalNumber
                FROM tblPatientMovement
                WHERE [category] = 'InterWardTransferIn' AND toWard = ?
                AND MovementDateTime BETWEEN ? AND ?
            ) AS TempTable", con)) // Added 'AS TempTable' alias
                            {
                                cmd.Parameters.AddWithValue("?", ward);
                                cmd.Parameters.AddWithValue("?", startDate);
                                cmd.Parameters.AddWithValue("?", endDate);
                                label6.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                            }
                            break;

                        case "InterWardTransferOut":
                            using (OleDbCommand cmd = new OleDbCommand(
                                @"SELECT COUNT(*) FROM (
                SELECT DISTINCT hospitalNumber
                FROM tblPatientMovement
                WHERE [category] = 'InterWardTransferOut' AND fromWard = ?
                AND MovementDateTime BETWEEN ? AND ?
            ) AS TempTable", con)) // Added 'AS TempTable' alias
                            {
                                cmd.Parameters.AddWithValue("?", ward);
                                cmd.Parameters.AddWithValue("?", startDate);
                                cmd.Parameters.AddWithValue("?", endDate);
                                label8.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                            }
                            break;

                        case "Death":
                            using (OleDbCommand cmd = new OleDbCommand(
                                @"SELECT COUNT(*) FROM tblPatientMaster 
              WHERE (currentWard = 'Mortuary' OR currentWard = 'Funeral Parlour') 
              AND isAdmitted = 'No'", con))
                            {
                                // Note: Since you hardcoded 'Mortuary' and 'Funeral Parlour', 
                                // no parameters are needed here.
                                label9.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error updating category: " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
        }



        private void button11_Click(object sender, EventArgs e)
        {
            textBox16.Clear();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string sql = @"INSERT INTO userAccounts 
               ([idNumber], [name], [surname], [password])
               VALUES (?, ?, ?, ?)";

                    OleDbCommand com = new OleDbCommand(sql, con);
                    com.Parameters.AddWithValue("?", textBox19.Text); // idNumber
                    com.Parameters.AddWithValue("?", textBox17.Text); // name
                    com.Parameters.AddWithValue("?", textBox18.Text); // surname
                    com.Parameters.AddWithValue("?", textBox20.Text); // password

                    com.ExecuteNonQuery();
                    MPHBSMS.LogActivity("ADMIN: Created user account for ID: " + textBox19.Text);
                    MessageBox.Show("Account successfully created\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    textBox17.Clear();
                    textBox18.Clear();
                    textBox19.Clear();
                    textBox20.Clear();
                    con.Close();
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    string sql = @"UPDATE userAccounts 
               SET [name]=?, [surname]=?, [password]=?
               WHERE [idNumber]=?";

                    OleDbCommand com = new OleDbCommand(sql, con);
                    com.Parameters.AddWithValue("?", textBox17.Text); // name
                    com.Parameters.AddWithValue("?", textBox18.Text); // surname
                    com.Parameters.AddWithValue("?", textBox20.Text); // password
                    com.Parameters.AddWithValue("?", textBox19.Text); // idNumber

                    com.ExecuteNonQuery();
                    MPHBSMS.LogActivity("ADMIN: Updated account details for ID: " + textBox19.Text);
                    MessageBox.Show("Account successfully Updated\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    textBox17.Clear();
                    textBox18.Clear();
                    textBox19.Clear();
                    textBox20.Clear();
                    con.Close();
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    string sql = @"DELETE FROM userAccounts WHERE [idNumber]=?";

                    OleDbCommand com = new OleDbCommand(sql, con);
                    com.Parameters.AddWithValue("?", textBox19.Text); // idNumber

                    com.ExecuteNonQuery();
                    MPHBSMS.LogActivity("SECURITY: Deleted user account ID: " + textBox19.Text);
                    con.Close();
                    MessageBox.Show("Account successfully Deleted\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {

            try
            {
                // 1. Clear the DataGridView before every new search
                dataGridView1.DataSource = null;
                if (dataGridView1.Rows.Count > 0) dataGridView1.Rows.Clear();

                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // Expanded SQL to support partial matches on all three identity fields
                    string sql = @"SELECT [idNumber], [name], [surname], [password]
                       FROM userAccounts
                       WHERE [idNumber] LIKE ? OR [name] LIKE ? OR [surname] LIKE ?";

                    using (OleDbCommand com = new OleDbCommand(sql, con))
                    {
                        // The % wildcards enable the "first three characters" search
                        string searchKey = "%" + textBox19.Text.Trim() + "%";
                        com.Parameters.AddWithValue("?", searchKey);
                        com.Parameters.AddWithValue("?", searchKey);
                        com.Parameters.AddWithValue("?", searchKey);

                        using (OleDbDataReader dr = com.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                // Map the first record to your Dashboard textboxes
                                textBox19.Text = dr["idNumber"].ToString();
                                textBox17.Text = dr["name"].ToString();
                                textBox18.Text = dr["surname"].ToString();
                                textBox20.Text = dr["password"].ToString();

                                // Close the reader so we can fill the DataGrid via Adapter
                                dr.Close();

                                // Fill the Grid with all matches found
                                OleDbDataAdapter dm = new OleDbDataAdapter(com);
                                DataTable dtt = new DataTable();
                                dm.Fill(dtt);
                                dataGridView1.DataSource = dtt;

                                MPHBSMS.LogActivity("ADMIN: Search completed for: " + textBox19.Text);
                            }
                            else
                            {
                                MessageBox.Show("No matching user profile found in Marondera Hospital records.", "Search Results");
                            }
                        }
                    }
                    con.Close();
                }
            }

            catch (Exception error)
            {
                MessageBox.Show("Error during search:\n" + error.Message, "System Error");
            }
        }
            

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                textBox17.Clear();
                textBox18.Clear();
                textBox19.Clear();
                textBox20.Clear();

                if (dataGridView1.DataSource != null)
                {
                    // This clears the data while keeping the column headers
                    ((DataTable)dataGridView1.DataSource).Rows.Clear();
                }
                else
                {
                    // Option B: If you added rows manually
                    dataGridView1.Rows.Clear();
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {

                    con.Open();

                    string myquaery = "SELECT * FROM userAccounts";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    // Inside the 'using' block after dm.Fill(dtt);
                    MPHBSMS.LogActivity("DATABASE SENSITIVE ACCESS: Loaded live user table to grid.");
                    dataGridView1.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // Local Validation
            if (string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox21.Text))
            {
                MessageBox.Show("Required Data Missing: Hospital Number and Movement ID must be filled.", "Marondera Provincial Hospital");
                return;
            }

            // Local Bed Day Calculation
            try
            {
                DateTime adm = DateTime.Parse(textBox13.Text);
                DateTime dis = DateTime.Parse(textBox10.Text);
                int totalDays = (dis.Date - adm.Date).Days;
                label17.Text = (totalDays <= 0 ? 1 : totalDays).ToString();
            }
            catch { label17.Text = "1"; }

            // Database Action
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string sql = "INSERT INTO tblPatientMovement (hospitalNumber, name, surname, gender, MovementDateTime, toWard, fromWard, category, enteredBy, movementID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.Parameters.AddWithValue("@1", textBox3.Text); cmd.Parameters.AddWithValue("@2", textBox4.Text);
                    cmd.Parameters.AddWithValue("@3", textBox5.Text); cmd.Parameters.AddWithValue("@4", textBox8.Text);
                    cmd.Parameters.AddWithValue("@5", textBox15.Text); cmd.Parameters.AddWithValue("@6", textBox6.Text);
                    cmd.Parameters.AddWithValue("@7", textBox7.Text); cmd.Parameters.AddWithValue("@8", textBox9.Text);
                    cmd.Parameters.AddWithValue("@9", textBox14.Text); cmd.Parameters.AddWithValue("@10", textBox21.Text);

                    cmd.ExecuteNonQuery();



                    con.Close();
                    MessageBox.Show("Record saved successfully.", "Marondera Provincial Hospital");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox21.Text))
            {
                MessageBox.Show("Select a record to delete.", "Marondera Provincial Hospital");
                return;
            }

            if (MessageBox.Show("Permanently delete this record?", "Marondera Provincial Hospital", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                    {
                        con.Open();
                        OleDbCommand cmd = new OleDbCommand("DELETE FROM tblPatientMovement WHERE movementID = ?", con);
                        cmd.Parameters.AddWithValue("@id", textBox21.Text);
                        cmd.ExecuteNonQuery();
                        con.Close();

                        label17.Text = "00";
                        MessageBox.Show("Record Deleted.", "Marondera Provincial Hospital");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR!!\n" + ex.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox21.Text) || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Please search for a record before attempting an update.", "Marondera Provincial Hospital");
                return;
            }

            try
            {
                DateTime adm = DateTime.Parse(textBox13.Text);
                DateTime dis = DateTime.Parse(textBox10.Text);
                label17.Text = (dis.Date - adm.Date).Days.ToString();
            }
            catch { }

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    // Update Master Table
                    OleDbCommand cmd1 = new OleDbCommand("UPDATE tblPatientMaster SET currentWard=?, dischargeDate=? WHERE hospitalNumber=?", con);
                    cmd1.Parameters.AddWithValue("@w", textBox12.Text);
                    cmd1.Parameters.AddWithValue("@d", textBox10.Text);
                    cmd1.Parameters.AddWithValue("@h", textBox3.Text);
                    cmd1.ExecuteNonQuery();

                    // Update Movement Table
                    OleDbCommand cmd2 = new OleDbCommand("UPDATE tblPatientMovement SET toWard=?, fromWard=? WHERE movementID=?", con);
                    cmd2.Parameters.AddWithValue("@t", textBox6.Text);
                    cmd2.Parameters.AddWithValue("@f", textBox7.Text);
                    cmd2.Parameters.AddWithValue("@m", textBox21.Text);
                    cmd2.ExecuteNonQuery();

                    con.Close();
                    MessageBox.Show("Update successful.", "Marondera Provincial Hospital");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

    // 1. Validation
    if (string.IsNullOrWhiteSpace(textBox3.Text))
    {
        MessageBox.Show("Please enter a Hospital Number.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }

    string targetHospNo = textBox3.Text.Trim();
    string currentUser = MPHBSMS.CurrentUser; 

    try
    {
        // 2. FIXED Path Handling - Creating the folder correctly
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string folderPath = Path.Combine(documentsPath, "Marondera Hospital Patient Files");
        
        // Ensure the main folder exists
        if (!Directory.Exists(folderPath)) 
        {
            Directory.CreateDirectory(folderPath);
        }

        // Remove any slashes or dots from the HospNo that might cause "Path not found"
        string cleanHospNo = string.Join("_", targetHospNo.Split(Path.GetInvalidFileNameChars()));
        
        // Simple filename structure to avoid nested folder errors
        string fileName = "History_" + cleanHospNo + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".pdf";
        string fullFilePath = Path.Combine(folderPath, fileName);

        // 3. Database Extraction (Same as before)
        DataTable historyData = new DataTable();
        using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
        {
            string query = "SELECT p.hospitalNumber, p.name, p.surname, p.admissionDate, " +
                           "m.category, m.MovementDateTime, m.fromWard, m.toWard, m.enteredBy " +
                           "FROM tblPatientMovement m " +
                           "INNER JOIN tblPatientMaster p ON m.hospitalNumber = p.hospitalNumber " +
                           "WHERE p.hospitalNumber = @hospNo " +
                           "ORDER BY m.MovementDateTime ASC";

            OleDbCommand cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@hospNo", targetHospNo);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(historyData);
        }

        if (historyData.Rows.Count == 0)
        {
            MessageBox.Show("No records found for: " + targetHospNo);
            return;
        }

        // 4. Generate PDF
        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate());
        using (FileStream fs = new FileStream(fullFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
            doc.Open();

            var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 14);
            var regFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 10);
            
            doc.Add(new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL - PATIENT AUDIT", titleFont));
            doc.Add(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("F"), regFont));
            doc.Add(new iTextSharp.text.Paragraph("Generated By: " + currentUser + "\n\n", regFont));

            iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(6);
            table.WidthPercentage = 100;
            string[] headers = { "Date", "Category", "From", "To", "Bed Days", "Entered By" };
            foreach (string h in headers) table.AddCell(new iTextSharp.text.Phrase(h, iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD)));

            int totalDays = 0;
            DateTime admission = Convert.ToDateTime(historyData.Rows[0]["admissionDate"]);

            foreach (DataRow row in historyData.Rows)
            {
                DateTime eventDate = Convert.ToDateTime(row["MovementDateTime"]);
                int days = (eventDate - admission).Days;
                if (days <= 0) days = 1;

                table.AddCell(eventDate.ToString("g"));
                table.AddCell(row["category"].ToString());
                table.AddCell(row["fromWard"].ToString());
                table.AddCell(row["toWard"].ToString());
                table.AddCell(days.ToString());
                table.AddCell(row["enteredBy"].ToString());
                totalDays = days;
            }
            doc.Add(table);
            doc.Add(new iTextSharp.text.Paragraph("\nTotal Cumulative Bed Days: " + totalDays));
            doc.Close();
        }

        // 5. Open the File
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullFilePath) { UseShellExecute = true });
        MPHBSMS.LogActivity("HISTORY VIEW: Generated and opened PDF for Patient " + targetHospNo);
    }
    catch (Exception ex)
    {
        MessageBox.Show("System Error: " + ex.Message);
    }

        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox8.Clear();
                textBox9.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox13.Clear();
                textBox14.Clear();
                textBox15.Clear();
                textBox21.Clear();
                label17.Text = "00";
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {

            DateTime startDate, endDate;

            // Validate Dates
            if (!DateTime.TryParse(textBox1.Text, out startDate) ||
                !DateTime.TryParse(textBox2.Text, out endDate))
            {
                button1.Enabled = false;
                return;
            }

            button1.Enabled = true;

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                try
                {
                    con.Open();

                    // 1. Total Admissions (All Wards)
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMaster WHERE isAdmitted = 'Yes'", con))
                    {
                        label5.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                    }

                    // 2. Total Discharges (All Wards)
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMaster WHERE isAdmitted = 'No'", con))
                    {
                        label7.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                    }

                    // 3. Total Inter-Ward Transfers IN (All Wards) within Date Range
                    using (OleDbCommand cmd = new OleDbCommand(
                        @"SELECT COUNT(*) FROM (
                SELECT DISTINCT hospitalNumber 
                FROM tblPatientMovement 
                WHERE [category] = 'InterWardTransferIn' 
                AND MovementDateTime BETWEEN ? AND ?
            ) AS TempTable", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label6.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                    }

                    // 4. Total Inter-Ward Transfers OUT (All Wards) within Date Range
                    using (OleDbCommand cmd = new OleDbCommand(
                        @"SELECT COUNT(*) FROM (
                SELECT DISTINCT hospitalNumber 
                FROM tblPatientMovement 
                WHERE [category] = 'InterWardTransferOut' 
                AND MovementDateTime BETWEEN ? AND ?
            ) AS TempTable", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label8.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                    }

                    // 5. Total Deaths (Mortuary/Funeral Parlour)
                    using (OleDbCommand cmd = new OleDbCommand(
                        @"SELECT COUNT(*) FROM tblPatientMaster 
              WHERE (currentWard = 'Mortuary' OR currentWard = 'Funeral Parlour') 
              AND isAdmitted = 'No'", con))
                    {
                        label9.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error calculating total statistics: " + ex.Message);
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string enteredBy = MPHBSMS.CurrentUser;

            // 1. Get the values from your textboxes/labels
            string start = textBox1.Text;
            string end = textBox2.Text;
            string user = enteredBy; // You can replace this with your actual user variable

            // 2. Call the method (this "triggers" the PDF generation)
            ExportSummaryToPDF(start, end, user);
        }

        // THIS METHOD MUST BE OUTSIDE THE BUTTON CLICK
        public void ExportSummaryToPDF(string startDateStr, string endDateStr, string currentUser)
        {
            try
            {
                // 1. Folder check/creation
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string folderName = "Marondera Provincial Hospital Bed Statistics Reports";
                string fullFolderPath = Path.Combine(documentsPath, folderName);

                if (!Directory.Exists(fullFolderPath))
                {
                    Directory.CreateDirectory(fullFolderPath);
                }

                // 2. Build Filename and Path with unique timestamp
                DateTime start = DateTime.Parse(startDateStr);
                DateTime end = DateTime.Parse(endDateStr);

                // Added HHmmss to ensure uniqueness even if generated in the same minute
                string desiredBaseName = string.Format("Report_{0}_to_{1}_{2}",
                                            start.ToString("yyyy-MM-dd"),
                                            end.ToString("yyyy-MM-dd"),
                                            DateTime.Now.ToString("HHmmss"));

                string fullFilePath = FilePathHelper.BuildSafeFilePath(fullFolderPath, desiredBaseName, ".pdf");



                // 3. Create Document
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30f, 30f, 30f, 30f);
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(fullFilePath, FileMode.Create));
                doc.Open();

                // --- HEADER ---
                iTextSharp.text.Paragraph header = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 16));
                header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                doc.Add(header);

                iTextSharp.text.Paragraph dept = new iTextSharp.text.Paragraph("DEPARTMENT: HEALTH INFORMATION", iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12));
                dept.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                doc.Add(dept);
                doc.Add(new iTextSharp.text.Paragraph("\n"));

                // --- METADATA ---
                doc.Add(new iTextSharp.text.Paragraph(string.Format("Generation Date: {0}", DateTime.Now.ToString("F"))));
                doc.Add(new iTextSharp.text.Paragraph(string.Format("Generated By: {0}", currentUser)));
                doc.Add(new iTextSharp.text.Paragraph(string.Format("Period: {0} to {1}", start.ToString("dd/MM/yyyy"), end.ToString("dd/MM/yyyy"))));

                string selectedWard = (comboBox2.SelectedItem != null) ? comboBox2.SelectedItem.ToString() : "All Wards";
                doc.Add(new iTextSharp.text.Paragraph(string.Format("Ward: {0}", selectedWard)));
                doc.Add(new iTextSharp.text.Paragraph("\n" + new string('-', 85) + "\n\n"));

                // --- TABLE ---
                iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(2);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 3f, 1.5f });

                iTextSharp.text.pdf.PdfPCell cell1 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Category", iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD)));
                cell1.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell1);

                iTextSharp.text.pdf.PdfPCell cell2 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Data Generated (Total)", iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD)));
                cell2.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell2);

                table.AddCell("Admissions"); table.AddCell(label5.Text);
                table.AddCell("Discharges"); table.AddCell(label7.Text);
                table.AddCell("Inter-Ward Transfers In"); table.AddCell(label6.Text);
                table.AddCell("Inter-Ward Transfers Out"); table.AddCell(label8.Text);
                table.AddCell("Deaths (Mortuary/Funeral)"); table.AddCell(label9.Text);

                doc.Add(table);

                // --- SIGNATURE ---
                doc.Add(new iTextSharp.text.Paragraph("\n\n\n"));
                iTextSharp.text.Paragraph sig = new iTextSharp.text.Paragraph("___________________________\nAuthorized Signature\nHealth Information Department");
                sig.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
                doc.Add(sig);

                doc.Close();
                System.Diagnostics.Process.Start(fullFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error: {0}", ex.Message), "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button23_Click(object sender, EventArgs e)
        {
            // 1. Check if a file is actually selected in the list
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("Please select a log file from the list to open.", "Marondera Provincial Hospital");
                return;
            }

            try
            {
                // 2. Get the filename and build the full path
                string fileName = listBox1.SelectedItem.ToString();
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                // 3. Verify the file exists on the disk before trying to open it
                if (File.Exists(filePath))
                {
                    // Opens the file using the system's default text editor (usually Notepad)
                    System.Diagnostics.Process.Start("notepad.exe", filePath);

                    // 4. Log the action so there is a record of who viewed the raw logs
                    MPHBSMS.LogActivity("RAW VIEW: Opened " + fileName + " in Notepad.");
                }
                else
                {
                    MessageBox.Show("The file " + fileName + " could not be found in the system folder.", "Marondera Provincial Hospital");
                }
            }
            catch (Exception ex)
            {
                // Concatenation used as requested
                MessageBox.Show("ERROR!!\n" + ex.Message, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            // Validation: Check ListBox and Search Box
            if (listBox1.SelectedItem == null || string.IsNullOrEmpty(textBox19.Text))
            {
                MessageBox.Show("Select a log file and enter an ID/Username to search.", "Marondera Provincial Hospital");
                return;
            }

            string inputSearch = textBox19.Text.Trim();
            string idMatch = "", nameMatch = "", surnameMatch = "";

            // 1. DATABASE LOOKUP: Find all identifiers for this person
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string sql = "SELECT idNumber, name, surname FROM userAccounts WHERE idNumber = ? OR name = ? OR surname = ?";
                    using (OleDbCommand cmd = new OleDbCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("?", inputSearch);
                        cmd.Parameters.AddWithValue("?", inputSearch);
                        cmd.Parameters.AddWithValue("?", inputSearch);

                        using (OleDbDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                idMatch = dr["idNumber"].ToString();
                                nameMatch = dr["name"].ToString();
                                surnameMatch = dr["surname"].ToString();
                            }
                            else { idMatch = inputSearch; }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Database identity sync failed: " + ex.Message); }

            // 2. FILE & CATEGORY PREPARATION
            string selectedFile = listBox1.SelectedItem.ToString();
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, selectedFile);
            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Unified_Extract_" + idMatch + ".txt");

            // Track totals based on your snapshot categories
            int countAdm = 0, countTransIn = 0, countDis = 0, countTransOut = 0, countDth = 0;

            try
            {
                string[] lines = File.ReadAllLines(sourcePath);
                System.Collections.Generic.List<string> matches = new System.Collections.Generic.List<string>();

                matches.Add("UNIFIED LOG REPORT FOR: " + nameMatch + " " + surnameMatch + " (ID: " + idMatch + ")");
                matches.Add("Generated on: " + DateTime.Now.ToString("F"));
                matches.Add("--------------------------------------------------");

                // 3. MULTI-KEY SCAN & CATEGORY COUNTING
                foreach (string singleLine in lines)
                {
                    bool isMatch = false;
                    if (!string.IsNullOrEmpty(idMatch) && singleLine.Contains(idMatch)) isMatch = true;
                    if (!string.IsNullOrEmpty(nameMatch) && singleLine.Contains(nameMatch)) isMatch = true;
                    if (!string.IsNullOrEmpty(surnameMatch) && singleLine.Contains(surnameMatch)) isMatch = true;

                    if (isMatch)
                    {
                        matches.Add(singleLine);

                        // Increment totals based on keywords in the action
                        string lowerLine = singleLine.ToLower();
                        if (lowerLine.Contains("admission")) countAdm++;
                        else if (lowerLine.Contains("transfer in")) countTransIn++;
                        else if (lowerLine.Contains("discharge")) countDis++;
                        else if (lowerLine.Contains("transfer out")) countTransOut++;
                        else if (lowerLine.Contains("death")) countDth++;
                    }
                }

                // 4. ADD SUMMARY TABLE AT THE BOTTOM (Mirroring Snapshot)
                matches.Add("\n--------------------------------------------------");
                matches.Add("STATISTICAL SUMMARY BY CATEGORY");
                matches.Add("--------------------------------------------------");
                matches.Add(string.Format("{0,-25} | {1}", "Total Admissions:", countAdm));
                matches.Add(string.Format("{0,-25} | {1}", "Total Transfers In:", countTransIn));
                matches.Add(string.Format("{0,-25} | {1}", "Total Discharges:", countDis));
                matches.Add(string.Format("{0,-25} | {1}", "Total Transfers Out:", countTransOut));
                matches.Add(string.Format("{0,-25} | {1}", "Total Deaths:", countDth));
                matches.Add("--------------------------------------------------");
                matches.Add("TOTAL ACTIONS: " + (countAdm + countTransIn + countDis + countTransOut + countDth));

                // 5. EXPORT AND OPEN
                if (matches.Count > 10) // Check if any actual data matches were found
                {
                    File.WriteAllLines(reportPath, matches.ToArray());
                    System.Diagnostics.Process.Start("notepad.exe", reportPath);
                    MPHBSMS.LogActivity("AUDIT: Exported unified log with Category Summary for " + idMatch);
                }
                else
                {
                    MessageBox.Show("No records found for this identity in the selected log.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("File Error: " + ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            try
            {
                // 1. Validation for ListBox and Search Input
                if (listBox1.SelectedItem == null || string.IsNullOrWhiteSpace(textBox19.Text))
                {
                    MessageBox.Show("Please select a log file and enter an ID/Name to filter.", "Marondera Provincial Hospital");
                    return;
                }

                string fileName = listBox1.SelectedItem.ToString();
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                string inputSearch = textBox19.Text.Trim();

                // 2. IDENTITY SYNC: Fetch all identifiers for the user
                string idMatch = "", nameMatch = "", surnameMatch = "";
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string sql = "SELECT idNumber, name, surname FROM userAccounts WHERE idNumber = ? OR name = ? OR surname = ?";
                    using (OleDbCommand cmd = new OleDbCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("?", inputSearch);
                        cmd.Parameters.AddWithValue("?", inputSearch);
                        cmd.Parameters.AddWithValue("?", inputSearch);
                        using (OleDbDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                idMatch = dr["idNumber"].ToString();
                                nameMatch = dr["name"].ToString();
                                surnameMatch = dr["surname"].ToString();
                            }
                            else { idMatch = inputSearch; }
                        }
                    }
                }

                // 3. PREPARE TRACKING AND GRID
                // Dictionary to store Category Name and Count
                System.Collections.Generic.Dictionary<string, int> categoryTotals = new System.Collections.Generic.Dictionary<string, int>();

                dataGridView1.Rows.Clear();
                if (dataGridView1.Columns.Count == 0)
                {
                    dataGridView1.Columns.Add("Date", "Date & Time");
                    dataGridView1.Columns.Add("User", "Identity Found");
                    dataGridView1.Columns.Add("Action", "Activity Logged");
                }

                // 4. PARSE AND FILTER LOGS
                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    bool isUserAction = false;
                    if (!string.IsNullOrEmpty(idMatch) && line.Contains(idMatch)) isUserAction = true;
                    if (!string.IsNullOrEmpty(nameMatch) && line.Contains(nameMatch)) isUserAction = true;
                    if (!string.IsNullOrEmpty(surnameMatch) && line.Contains(surnameMatch)) isUserAction = true;

                    if (isUserAction)
                    {
                        string[] parts = line.Split('|');
                        if (parts.Length >= 3)
                        {
                            string time = parts[0].Replace("[", "").Replace("]", "").Trim();
                            string user = parts[1].Replace("User:", "").Trim();
                            string action = parts[2].Replace("Action:", "").Trim();

                            dataGridView1.Rows.Add(time, user, action);

                            // --- NEW: Category Calculation ---
                            // We identify the category by looking at the first word of the action (e.g., "ADMISSION", "VIEW")
                            string category = action.Split(' ')[0].ToUpper().Replace(":", "");

                            if (categoryTotals.ContainsKey(category))
                                categoryTotals[category]++;
                            else
                                categoryTotals[category] = 1;
                        }
                    }
                }

                // 5. APPEND TOTALS TO GRID (OR SHOW IN MESSAGE)
                if (categoryTotals.Count > 0)
                {
                    // Add a separator row for visual clarity
                    dataGridView1.Rows.Add("---", "--- SUMMARY ---", "---");

                    foreach (var entry in categoryTotals)
                    {
                        // Add a row showing: Category Name | Total | Count
                        dataGridView1.Rows.Add("TOTAL", entry.Key, entry.Value.ToString() + " Actions performed");
                    }

                    // Highlight the summary rows (optional)
                    int lastRowIndex = dataGridView1.Rows.Count - 1;
                    dataGridView1.Rows[lastRowIndex].DefaultCellStyle.BackColor = Color.LightGray;
                }

                // 6. Audit Logging
                MPHBSMS.LogActivity("AUDIT FILTER: Viewed unified logs for " + (nameMatch ?? inputSearch));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading unified log: " + ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {

        }

        private void button25_Click(object sender, EventArgs e)
        {
            // Calling the method from your new class
            BackupManager.ExecuteAdvancedBackup();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            // Safety check: Restoring data is permanent
            DialogResult result = MessageBox.Show("This will overwrite all current patient records with the backup. Continue?",
                                                  "Marondera Hospital Security", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                BackupManager.RestoreSystemFromBackup();
            }
        }
       
        private void UpdateBackupStatusLabel()
        {
            try
            {
                // 1. Locate the status file in the application folder
                string statusPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backup_status.txt");

                if (File.Exists(statusPath))
                {
                    // 2. Read the date of the last successful backup
                    string lastDateStr = File.ReadAllText(statusPath);
                    DateTime lastBackupDate = DateTime.Parse(lastDateStr);

                    lblLastBackup.Text = "Last Backup: " + lastBackupDate.ToString("dd MMM yyyy HH:mm");

                    // 3. Calculate if it's been more than 48 hours
                    TimeSpan timeSinceBackup = DateTime.Now - lastBackupDate;

                    if (timeSinceBackup.TotalHours > 48)
                    {
                        lblLastBackup.ForeColor = Color.Red;
                        // Immediate alert for the hospital administrator
                        MessageBox.Show("CRITICAL: System backup is " + (int)timeSinceBackup.TotalDays + " days overdue!",
                                        "Marondera Hospital Security Alert",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        lblLastBackup.ForeColor = Color.Green; // Backup is healthy
                    }
                }
                else
                {
                    lblLastBackup.Text = "Last Backup: Never";
                    lblLastBackup.ForeColor = Color.Red;
                }
            }
            catch
            {
                lblLastBackup.Text = "Status: Check Failed";
            }
        }
    }

}      
       