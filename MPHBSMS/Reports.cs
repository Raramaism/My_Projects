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
            // 1. INITIALIZE VARIABLES (Fixes 'unassigned local variable' error)
            DateTime start = DateTime.MinValue;
            DateTime end = DateTime.MinValue;
            string currentUser = MPHBSMS.CurrentUser;

            // 2. VALIDATE DATES
            if (!DateTime.TryParse(textBox1.Text, out start) || !DateTime.TryParse(textBox2.Text, out end))
            {
                MessageBox.Show("Please ensure valid dates are entered in the textboxes before exporting.");
                return;
            }

            try
            {
                // 3. FOLDER & FILE SETUP
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string folderPath = Path.Combine(documentsPath, "Marondera Provincial Hospital Bed Statistics Reports", "Detailed Reports");
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                string fileName = string.Format("Detailed_Stats_{0}_{1}.pdf",
                                    start.ToString("yyyyMMdd"),
                                    DateTime.Now.ToString("HHmmss"));
                string fullPath = Path.Combine(folderPath, fileName);

                // 4. DATABASE EXTRACTION (Flexible Search Logic)
                string selectedWard = (comboBox2.SelectedItem != null) ? comboBox2.SelectedItem.ToString() : "";
                string selectedCat = (comboBox1.SelectedItem != null) ? comboBox1.SelectedItem.ToString() : "";

                DataTable reportData = new DataTable();
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // Base Query
                    string sql = "SELECT [MovementDateTime], [hospitalNumber], [name], [surname], [category], [fromWard], [toWard] " +
                                 "FROM [tblPatientMovement] " +
                                 "WHERE ([MovementDateTime] BETWEEN ? AND ?) ";

                    // Optional Ward Filter
                    if (!string.IsNullOrEmpty(selectedWard))
                    {
                        sql += "AND ([fromWard] = ? OR [toWard] = ?) ";
                    }
                    // Optional Category Filter
                    if (!string.IsNullOrEmpty(selectedCat))
                    {
                        sql += "AND ([category] = ?) ";
                    }

                    sql += "ORDER BY [MovementDateTime] DESC";

                    using (OleDbCommand cmd = new OleDbCommand(sql, con))
                    {
                        // Add parameters in the exact order of the '?' marks
                        cmd.Parameters.AddWithValue("?", start);
                        cmd.Parameters.AddWithValue("?", end);

                        if (!string.IsNullOrEmpty(selectedWard))
                        {
                            cmd.Parameters.AddWithValue("?", selectedWard);
                            cmd.Parameters.AddWithValue("?", selectedWard);
                        }
                        if (!string.IsNullOrEmpty(selectedCat))
                        {
                            cmd.Parameters.AddWithValue("?", selectedCat);
                        }

                        new OleDbDataAdapter(cmd).Fill(reportData);
                    }
                }

                // 5. PDF GENERATION
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30f, 30f, 30f, 30f);
                using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                {
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // Logo Fix (Centers logo at top)
                    System.Drawing.Image resImage = Properties.Resources.logo1;
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(resImage, System.Drawing.Imaging.ImageFormat.Png);
                    logo.ScaleToFit(75f, 75f);
                    logo.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                    doc.Add(logo);

                    var titleFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 16);
                    var subFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 11);
                    var regFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 9);

                    iTextSharp.text.Paragraph head = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", titleFont);
                    head.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    doc.Add(head);

                    iTextSharp.text.Paragraph subHead = new iTextSharp.text.Paragraph("DETAILED PATIENT MOVEMENT & BED STATISTICS", subFont);
                    subHead.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    doc.Add(subHead);

                    doc.Add(new iTextSharp.text.Paragraph("Period: " + start.ToString("dd/MM/yyyy") + " to " + end.ToString("dd/MM/yyyy") + " | Generated By: " + currentUser, regFont));
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // Summary Totals Table
                    iTextSharp.text.pdf.PdfPTable sumTable = new iTextSharp.text.pdf.PdfPTable(2);
                    sumTable.WidthPercentage = 40;
                    sumTable.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

                    sumTable.AddCell(new iTextSharp.text.Phrase("Total Records Found:", regFont));
                    sumTable.AddCell(new iTextSharp.text.Phrase(reportData.Rows.Count.ToString(), regFont));
                    doc.Add(sumTable);
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // Detailed Log Table
                    iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(7);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 2f, 1.2f, 2f, 2f, 1.5f, 1.5f, 1.5f });

                    string[] headers = { "Date/Time", "Hosp #", "Name", "Surname", "Category", "From", "To" };
                    foreach (string h in headers)
                    {
                        iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(h, subFont));
                        cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                        table.AddCell(cell);
                    }

                    foreach (DataRow row in reportData.Rows)
                    {
                        table.AddCell(new iTextSharp.text.Phrase(Convert.ToDateTime(row["MovementDateTime"]).ToString("dd/MM/yyyy HH:mm"), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["hospitalNumber"].ToString(), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["name"].ToString(), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["surname"].ToString(), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["category"].ToString(), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["fromWard"].ToString(), regFont));
                        table.AddCell(new iTextSharp.text.Phrase(row["toWard"].ToString(), regFont));
                    }
                    doc.Add(table);
                    doc.Close();
                }

                // 6. RESET UI ELEMENTS
                textBox1.Clear();
                textBox2.Clear();
                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
                label5.Text = "0"; label6.Text = "0"; label7.Text = "0"; label8.Text = "0"; label9.Text = "0";

                // 7. OPEN FILE
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating detailed report: " + ex.Message);
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

            string targetHospNo = textBox3.Text.Trim().ToUpper();
            string currentUser = MPHBSMS.CurrentUser;

            try
            {
                // 2. FOLDER AND SUBFOLDER HANDLING
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string mainFolderPath = Path.Combine(documentsPath, "Marondera Hospital Patient Files");
                string subFolderPath = Path.Combine(mainFolderPath, "Patient History Report");

                if (!Directory.Exists(mainFolderPath)) Directory.CreateDirectory(mainFolderPath);
                if (!Directory.Exists(subFolderPath)) Directory.CreateDirectory(subFolderPath);

                string cleanHospNo = string.Join("_", targetHospNo.Split(Path.GetInvalidFileNameChars()));
                string fileName = "FullAudit_" + cleanHospNo + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".pdf";
                string fullFilePath = Path.Combine(subFolderPath, fileName);

                // 3. Database Extraction
                DataTable masterData = new DataTable();
                DataTable historyData = new DataTable();

                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string masterQuery = "SELECT * FROM [tblPatientMaster] WHERE [hospitalNumber] = ?";
                    using (OleDbCommand cmd = new OleDbCommand(masterQuery, con))
                    {
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value = targetHospNo;
                        new OleDbDataAdapter(cmd).Fill(masterData);
                    }

                    string historyQuery = "SELECT * FROM [tblPatientMovement] WHERE [hospitalNumber] = ? ORDER BY [MovementDateTime] ASC";
                    using (OleDbCommand cmd = new OleDbCommand(historyQuery, con))
                    {
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value = targetHospNo;
                        new OleDbDataAdapter(cmd).Fill(historyData);
                    }
                }

                if (masterData.Rows.Count == 0)
                {
                    MessageBox.Show("No record found for Hospital Number: " + targetHospNo);
                    return;
                }

                // 4. PDF Generation (Landscape)
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate());
                using (FileStream fs = new FileStream(fullFilePath, FileMode.Create))
                {
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // --- WATERMARK ---
                    iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContentUnder;
                    iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA_BOLD, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 55);
                    cb.SetColorFill(new iTextSharp.text.BaseColor(230, 230, 230));
                    cb.ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, "OFFICIAL PATIENT AUDIT", doc.PageSize.Width / 2, doc.PageSize.Height / 2, 35);
                    cb.EndText();

                    // --- LOGO LOGIC (Using Resources) ---
                    // By adding the logo directly to 'doc' without AbsolutePosition, 
                    // the text will automatically start BELOW it.
                    System.Drawing.Image resImage = Properties.Resources.logo1;
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(resImage, System.Drawing.Imaging.ImageFormat.Png);
                    logo.ScaleToFit(70f, 70f);
                    logo.Alignment = iTextSharp.text.Image.ALIGN_CENTER; // Center the logo
                    doc.Add(logo);

                    // Fonts
                    var titleFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 16);
                    var regFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 10);
                    var tableHeaderFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 9);
                    var smallFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 8);

                    // Header Title (Now appears below image)
                    iTextSharp.text.Paragraph header = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL - FULL AUDIT REPORT", titleFont);
                    header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    doc.Add(header);

                    doc.Add(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("F"), regFont));
                    doc.Add(new iTextSharp.text.Paragraph("Generated By: " + currentUser, regFont));
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // SECTION 1: MASTER DATA
                    DataRow mRow = masterData.Rows[0];
                    iTextSharp.text.pdf.PdfPTable masterTable = new iTextSharp.text.pdf.PdfPTable(4);
                    masterTable.WidthPercentage = 100;
                    masterTable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    masterTable.AddCell(new iTextSharp.text.Phrase("Hospital No:", tableHeaderFont));
                    masterTable.AddCell(new iTextSharp.text.Phrase(mRow["hospitalNumber"].ToString(), regFont));
                    masterTable.AddCell(new iTextSharp.text.Phrase("Patient Name:", tableHeaderFont));
                    masterTable.AddCell(new iTextSharp.text.Phrase(mRow["name"].ToString() + " " + mRow["surname"].ToString(), regFont));
                    doc.Add(masterTable);
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // SECTION 2: HISTORY TABLE (10 Columns)
                    iTextSharp.text.pdf.PdfPTable historyTable = new iTextSharp.text.pdf.PdfPTable(10);
                    historyTable.WidthPercentage = 100;
                    string[] headers = { "Hosp No", "Name", "Surname", "Gender", "Movement Date", "To Ward", "From Ward", "Category", "Entered By", "ID" };

                    foreach (string h in headers)
                    {
                        iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(h, tableHeaderFont));
                        cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                        cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        historyTable.AddCell(cell);
                    }

                    foreach (DataRow row in historyData.Rows)
                    {
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["hospitalNumber"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["name"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["surname"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["gender"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["MovementDateTime"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["toWard"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["fromWard"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["category"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["enteredBy"].ToString(), smallFont));
                        historyTable.AddCell(new iTextSharp.text.Phrase(row["movementID"].ToString(), smallFont));
                    }
                    doc.Add(historyTable);

                    // SECTION 3: BED DAYS
                    doc.Add(new iTextSharp.text.Paragraph("\n" + new string('-', 140), regFont));

                    DateTime admDate;
                    if (DateTime.TryParse(mRow["admissionDate"].ToString(), out admDate))
                    {
                        DateTime endDate = DateTime.Now;
                        if (mRow["isAdmitted"].ToString().ToLower() == "no")
                        {
                            DateTime.TryParse(mRow["dischargeDate"].ToString(), out endDate);
                        }
                        TimeSpan span = endDate - admDate;
                        int totalBedDays = Math.Max(1, (int)Math.Ceiling(span.TotalDays));
                        doc.Add(new iTextSharp.text.Paragraph("TOTAL CUMULATIVE BED DAYS: " + totalBedDays, titleFont));
                    }

                    // SIGNATURE
                    doc.Add(new iTextSharp.text.Paragraph("\n\n\n___________________________\nAuthorized Signature\nHealth Information Department", smallFont));

                    doc.Close();
                }

                // Open file
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullFilePath) { UseShellExecute = true });
                MPHBSMS.LogActivity("AUDIT REPORT: Generated for Hospital No " + targetHospNo);
            }
            catch (Exception ex) { MessageBox.Show("Critical Error: " + ex.Message); }
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

                    // 1. Total Admissions (Within Date Range)
                    // Uses tblPatientMovement to find when category was set to 'Admission'
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMovement WHERE [category] = 'Admission' AND MovementDateTime BETWEEN ? AND ?", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label5.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 2. Total Discharges (Within Date Range)
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMovement WHERE [category] = 'Discharge' AND MovementDateTime BETWEEN ? AND ?", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label7.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 3. Total Inter-Ward Transfers IN (Within Date Range)
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMovement WHERE [category] = 'InterWardTransferIn' AND MovementDateTime BETWEEN ? AND ?", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label6.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 4. Total Inter-Ward Transfers OUT (Within Date Range)
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMovement WHERE [category] = 'InterWardTransferOut' AND MovementDateTime BETWEEN ? AND ?", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label8.Text = cmd.ExecuteScalar().ToString();
                    }

                    // 5. Total Deaths (Based on Movement Records within range)
                    // This is more accurate than checking currentWard in PatientMaster
                    using (OleDbCommand cmd = new OleDbCommand(
                        "SELECT COUNT(*) FROM tblPatientMovement WHERE [category] = 'Death' AND MovementDateTime BETWEEN ? AND ?", con))
                    {
                        cmd.Parameters.AddWithValue("?", startDate);
                        cmd.Parameters.AddWithValue("?", endDate);
                        label9.Text = cmd.ExecuteScalar().ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error calculating total statistics: " + ex.Message); //
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string enteredBy = MPHBSMS.CurrentUser;

            // 1. Get values and call method
            string start = textBox1.Text;
            string end = textBox2.Text;
            string user = enteredBy;

            ExportSummaryToPDF(start, end, user);
        }
public void ExportSummaryToPDF(string startDateStr, string endDateStr, string currentUser)
{
    try
    {
        // 1. FOLDER HANDLING
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string mainFolderName = "Marondera Provincial Hospital Bed Statistics Reports";
        string subFolderName = "General Statistics";
        
        string mainPath = Path.Combine(documentsPath, mainFolderName);
        string subPath = Path.Combine(mainPath, subFolderName);

        if (!Directory.Exists(mainPath)) Directory.CreateDirectory(mainPath);
        if (!Directory.Exists(subPath)) Directory.CreateDirectory(subPath);

        // 2. BUILD FILENAME
        DateTime start = DateTime.Parse(startDateStr);
        DateTime end = DateTime.Parse(endDateStr);

        string desiredBaseName = string.Format("Stats_{0}_to_{1}_{2}",
                                    start.ToString("yyyy-MM-dd"),
                                    end.ToString("yyyy-MM-dd"),
                                    DateTime.Now.ToString("HHmmss"));

        string fullFilePath = Path.Combine(subPath, desiredBaseName + ".pdf");

        // 3. CREATE DOCUMENT
        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30f, 30f, 30f, 30f);
        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(fullFilePath, FileMode.Create));
        
        doc.Open();

        // --- WATERMARK ---
        iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContentUnder;
        iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA_BOLD, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);
        cb.BeginText();
        cb.SetFontAndSize(bf, 45);
        cb.SetColorFill(new iTextSharp.text.BaseColor(235, 235, 235)); 
        cb.ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, "OFFICIAL BED STATISTICS REPORT", doc.PageSize.Width / 2, doc.PageSize.Height / 2, 45);
        cb.EndText();

        // --- LOGO (Fix: Internal Resource to avoid missing file errors) ---
        // Pulling directly from Resources.resx as seen in your project
        System.Drawing.Image resImage = Properties.Resources.logo1;
        iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(resImage, System.Drawing.Imaging.ImageFormat.Png);
        logo.ScaleToFit(80f, 80f);
        logo.Alignment = iTextSharp.text.Image.ALIGN_CENTER; 
        doc.Add(logo); // Fix: Adding directly prevents text overlap

        // --- HEADER ---
        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 16);
        var subFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
        
        iTextSharp.text.Paragraph header = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", titleFont);
        header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
        doc.Add(header);

        iTextSharp.text.Paragraph dept = new iTextSharp.text.Paragraph("DEPARTMENT: HEALTH INFORMATION", subFont);
        dept.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
        doc.Add(dept);
        doc.Add(new iTextSharp.text.Paragraph("\n"));

        // --- METADATA (No $ Interpolation used here) ---
        doc.Add(new iTextSharp.text.Paragraph(string.Format("Generation Date: {0}", DateTime.Now.ToString("F"))));
        doc.Add(new iTextSharp.text.Paragraph(string.Format("Generated By: {0}", currentUser)));
        doc.Add(new iTextSharp.text.Paragraph(string.Format("Reporting Period: {0} to {1}", start.ToString("dd/MM/yyyy"), end.ToString("dd/MM/yyyy"))));

        string selectedWard = (comboBox2.SelectedItem != null) ? comboBox2.SelectedItem.ToString() : "Hospital-Wide";
        doc.Add(new iTextSharp.text.Paragraph(string.Format("Ward Scope: {0}", selectedWard)));
        doc.Add(new iTextSharp.text.Paragraph("\n" + new string('-', 85) + "\n\n"));

        // --- DATA TABLE (Populated with current calculated totals) ---
        iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(2);
        table.WidthPercentage = 100;
        table.SetWidths(new float[] { 3f, 1.5f });

        var hFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 11);
        
        // Header Cells
        iTextSharp.text.pdf.PdfPCell hCell1 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Category", hFont));
        hCell1.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
        hCell1.Padding = 6;
        table.AddCell(hCell1);

        iTextSharp.text.pdf.PdfPCell hCell2 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Total Count", hFont));
        hCell2.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
        hCell2.Padding = 6;
        table.AddCell(hCell2);

        // Populate from labels (Ensure statistics are updated first)
        table.AddCell("Total Admissions"); table.AddCell(label5.Text);
        table.AddCell("Total Discharges"); table.AddCell(label7.Text);
        table.AddCell("Total Inter-Ward Transfers In"); table.AddCell(label6.Text);
        table.AddCell("Total Inter-Ward Transfers Out"); table.AddCell(label8.Text);
        table.AddCell("Total Deaths Recorded"); table.AddCell(label9.Text);

        doc.Add(table);

        // --- SIGNATURE SECTION ---
        doc.Add(new iTextSharp.text.Paragraph("\n\n\n"));
        iTextSharp.text.Paragraph sig = new iTextSharp.text.Paragraph("___________________________\nAuthorized Signature\nHealth Information Department");
        sig.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
        doc.Add(sig);

        doc.Close();

        // 5. Open file automatically
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullFilePath) { UseShellExecute = true });
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
            try
            {
                // 1. DATA VALIDATION & FOLDER SETUP
                DateTime start, end;
                if (!DateTime.TryParse(textBox1.Text, out start) || !DateTime.TryParse(textBox2.Text, out end))
                {
                    MessageBox.Show("Please ensure valid dates are entered in the textboxes before exporting.");
                    return;
                }

                string currentUser = MPHBSMS.CurrentUser;
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string mainPath = Path.Combine(documentsPath, "Marondera Provincial Hospital Bed Statistics Reports");
                string subPath = Path.Combine(mainPath, "General Statistics");

                if (!Directory.Exists(mainPath)) Directory.CreateDirectory(mainPath);
                if (!Directory.Exists(subPath)) Directory.CreateDirectory(subPath);

                // Build Filename using string.Format (No $ interpolation)
                string fileName = string.Format("Stats_{0}_to_{1}_{2}.pdf",
                                    start.ToString("yyyy-MM-dd"),
                                    end.ToString("yyyy-MM-dd"),
                                    DateTime.Now.ToString("HHmmss"));

                string fullFilePath = Path.Combine(subPath, fileName);

                // 2. DATABASE CALCULATION LOGIC
                // We calculate fresh totals directly before generating the PDF to ensure accuracy
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    // Using tblPatientMovement for accurate historical reporting
                    string[] categories = { "Admission", "Discharge", "InterWardTransferIn", "InterWardTransferOut", "Death" };
                    string[] counts = new string[5];

                    for (int i = 0; i < categories.Length; i++)
                    {
                        // Using brackets [] for Access SQL safety and positional ? parameters
                        string sql = "SELECT COUNT(*) FROM [tblPatientMovement] WHERE [category] = ? AND [MovementDateTime] BETWEEN ? AND ?";
                        using (OleDbCommand cmd = new OleDbCommand(sql, con))
                        {
                            cmd.Parameters.AddWithValue("?", categories[i]);
                            cmd.Parameters.AddWithValue("?", start);
                            cmd.Parameters.AddWithValue("?", end);
                            counts[i] = cmd.ExecuteScalar().ToString();
                        }
                    }

                    // Sync form labels with the new data
                    label5.Text = counts[0]; // Admissions
                    label7.Text = counts[1]; // Discharges
                    label6.Text = counts[2]; // Transfer In
                    label8.Text = counts[3]; // Transfer Out
                    label9.Text = counts[4]; // Deaths
                }

                // 3. PDF GENERATION
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30f, 30f, 30f, 30f);
                using (FileStream fs = new FileStream(fullFilePath, FileMode.Create))
                {
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // --- WATERMARK ---
                    iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContentUnder;
                    iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA_BOLD, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 45);
                    cb.SetColorFill(new iTextSharp.text.BaseColor(240, 240, 240));
                    cb.ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, "OFFICIAL BED STATISTICS", doc.PageSize.Width / 2, doc.PageSize.Height / 2, 45);
                    cb.EndText();

                    // --- LOGO (Alignment centered avoids text overlap) ---
                    System.Drawing.Image resImage = Properties.Resources.logo1;
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(resImage, System.Drawing.Imaging.ImageFormat.Png);
                    logo.ScaleToFit(85f, 85f);
                    logo.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                    doc.Add(logo);

                    // --- FONTS ---
                    var titleFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 16);
                    var subFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 12);
                    var regFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 10);

                    // --- HEADERS ---
                    iTextSharp.text.Paragraph h1 = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", titleFont);
                    h1.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    doc.Add(h1);

                    iTextSharp.text.Paragraph h2 = new iTextSharp.text.Paragraph("HEALTH INFORMATION DEPARTMENT - STATISTICS SUMMARY", subFont);
                    h2.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                    doc.Add(h2);
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // --- METADATA (String concatenation only) ---
                    doc.Add(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("F"), regFont));
                    doc.Add(new iTextSharp.text.Paragraph("Generated By: " + currentUser, regFont));
                    doc.Add(new iTextSharp.text.Paragraph("Reporting Period: " + start.ToString("dd/MM/yyyy") + " to " + end.ToString("dd/MM/yyyy"), regFont));
                    doc.Add(new iTextSharp.text.Paragraph("\n" + new string('-', 90) + "\n\n", regFont));

                    // --- DATA TABLE ---
                    iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(2);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 3.5f, 1.5f });

                    // Table Header Cells
                    iTextSharp.text.pdf.PdfPCell c1 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Statistical Category", subFont));
                    c1.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    c1.Padding = 8;
                    table.AddCell(c1);

                    iTextSharp.text.pdf.PdfPCell c2 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Total Count", subFont));
                    c2.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    c2.Padding = 8;
                    table.AddCell(c2);

                    // Adding data rows
                    table.AddCell("Total Admissions"); table.AddCell(label5.Text);
                    table.AddCell("Total Discharges"); table.AddCell(label7.Text);
                    table.AddCell("Total Inter-Ward Transfers In"); table.AddCell(label6.Text);
                    table.AddCell("Total Inter-Ward Transfers Out"); table.AddCell(label8.Text);
                    table.AddCell("Total Deaths Recorded"); table.AddCell(label9.Text);

                    doc.Add(table);

                    // --- SIGNATURE SECTION ---
                    doc.Add(new iTextSharp.text.Paragraph("\n\n\n\n"));
                    iTextSharp.text.Paragraph sig = new iTextSharp.text.Paragraph("___________________________\nAuthorized Signature\nRecords Office", regFont);
                    sig.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
                    doc.Add(sig);

                    doc.Close();
                }

                // 4. AUTO-OPEN THE FINAL REPORT
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullFilePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                // Concatenation used for error message
                MessageBox.Show("Export Error: " + ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            try
            {
                // 1. Validation
                if (listBox1.SelectedItem == null || string.IsNullOrWhiteSpace(textBox19.Text))
                {
                    MessageBox.Show("Please select a log file and enter an ID/Name to filter.", "Marondera Provincial Hospital");
                    return;
                }

                string fileName = listBox1.SelectedItem.ToString();
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                string inputSearch = textBox19.Text.Trim();

                // 2. Identity Sync (Database Lookup)
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

                // 3. Prepare PDF Path
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string subFolder = Path.Combine(documentsPath, "Marondera Hospital Patient Files", "User Audit Reports");
                if (!Directory.Exists(subFolder)) Directory.CreateDirectory(subFolder);
                string pdfName = "AuditLog_" + idMatch + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".pdf";
                string fullPath = Path.Combine(subFolder, pdfName);

                // 4. PDF Generation
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate()); // Rotate for better space
                using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                {
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // --- Watermark ---
                    iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContentUnder;
                    iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA_BOLD, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 50);
                    cb.SetColorFill(new iTextSharp.text.BaseColor(240, 240, 240));
                    cb.ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, "SECURITY AUDIT LOG", doc.PageSize.Width / 2, doc.PageSize.Height / 2, 45);
                    cb.EndText();

                    // --- Logo (logo.jpg) ---
                    string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.jpg");
                    if (File.Exists(logoPath))
                    {
                        iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(logoPath);
                        logo.ScaleToFit(70f, 70f);
                        logo.SetAbsolutePosition(35f, doc.PageSize.Height - 85f);
                        doc.Add(logo);
                    }

                    // Fonts
                    var titleFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 16);
                    var regFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 10);
                    var tableHeaderFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 10);

                    doc.Add(new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", titleFont));
                    doc.Add(new iTextSharp.text.Paragraph("USER ACTIVITY AUDIT REPORT", iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 12)));
                    doc.Add(new iTextSharp.text.Paragraph("Target User: " + nameMatch + " " + surnameMatch + " (" + idMatch + ")", regFont));
                    doc.Add(new iTextSharp.text.Paragraph("Log File: " + fileName, regFont));
                    doc.Add(new iTextSharp.text.Paragraph("Generated On: " + DateTime.Now.ToString("F"), regFont));
                    doc.Add(new iTextSharp.text.Paragraph("\n"));

                    // 5. Build Main Activity Table
                    iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(3);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 2.5f, 2.5f, 5f });

                    table.AddCell(new iTextSharp.text.Phrase("Date/Time", tableHeaderFont));
                    table.AddCell(new iTextSharp.text.Phrase("User Identity", tableHeaderFont));
                    table.AddCell(new iTextSharp.text.Phrase("Action Performed", tableHeaderFont));

                    System.Collections.Generic.Dictionary<string, int> categoryTotals = new System.Collections.Generic.Dictionary<string, int>();
                    string[] lines = File.ReadAllLines(filePath);

                    foreach (string line in lines)
                    {
                        bool isUserAction = (line.Contains(idMatch) || (nameMatch != "" && line.Contains(nameMatch)) || (surnameMatch != "" && line.Contains(surnameMatch)));

                        if (isUserAction)
                        {
                            string[] parts = line.Split('|');
                            if (parts.Length >= 3)
                            {
                                string time = parts[0].Replace("[", "").Replace("]", "").Trim();
                                string user = parts[1].Replace("User:", "").Trim();
                                string action = parts[2].Replace("Action:", "").Trim();

                                table.AddCell(new iTextSharp.text.Phrase(time, regFont));
                                table.AddCell(new iTextSharp.text.Phrase(user, regFont));
                                table.AddCell(new iTextSharp.text.Phrase(action, regFont));

                                // Category Calculation
                                string category = action.Split(' ')[0].ToUpper().Replace(":", "");
                                if (categoryTotals.ContainsKey(category)) categoryTotals[category]++;
                                else categoryTotals[category] = 1;
                            }
                        }
                    }
                    doc.Add(table);

                    // 6. Summary Section (The Totals)
                    doc.Add(new iTextSharp.text.Paragraph("\nACTIVITY SUMMARY", tableHeaderFont));
                    iTextSharp.text.pdf.PdfPTable summaryTable = new iTextSharp.text.pdf.PdfPTable(2);
                    summaryTable.WidthPercentage = 40;
                    summaryTable.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

                    foreach (var entry in categoryTotals)
                    {
                        summaryTable.AddCell(new iTextSharp.text.Phrase(entry.Key, regFont));
                        summaryTable.AddCell(new iTextSharp.text.Phrase(entry.Value.ToString() + " Times", regFont));
                    }
                    doc.Add(summaryTable);

                    doc.Close();
                }

                // Open PDF
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullPath) { UseShellExecute = true });
                MPHBSMS.LogActivity("AUDIT EXPORT: Generated PDF Log Audit for " + idMatch);
            }
            catch (Exception ex) { MessageBox.Show("Critical Error: " + ex.Message); }
        }

        private void button24_Click(object sender, EventArgs e)
        {
    try
    {
        // 1. VALIDATION & DIRECTORY SETUP
        DateTime start, end;
        if (!DateTime.TryParse(textBox1.Text, out start) || !DateTime.TryParse(textBox2.Text, out end))
        {
            MessageBox.Show("Please ensure valid dates are entered before exporting.");
            return;
        }

        string currentUser = MPHBSMS.CurrentUser; 
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string folderPath = Path.Combine(documentsPath, "Marondera Provincial Hospital Bed Statistics Reports", "Detailed Reports");
        if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

        string fileName = string.Format("Detailed_Stats_{0}.pdf", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        string fullPath = Path.Combine(folderPath, fileName);

        // 2. DATABASE EXTRACTION (Pulling Raw Movement Data)
        DataTable movementData = new DataTable();
        using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
        {
            con.Open();
            // Select detailed movement records within the date range
            string sql = "SELECT [MovementDateTime], [hospitalNumber], [name], [surname], [category], [fromWard], [toWard] " +
                         "FROM [tblPatientMovement] " +
                         "WHERE [MovementDateTime] BETWEEN ? AND ? " +
                         "ORDER BY [MovementDateTime] DESC";
            
            using (OleDbCommand cmd = new OleDbCommand(sql, con))
            {
                cmd.Parameters.AddWithValue("?", start);
                cmd.Parameters.AddWithValue("?", end);
                new OleDbDataAdapter(cmd).Fill(movementData);
            }
        }

        // 3. PDF GENERATION
        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30f, 30f, 30f, 30f); // Landscape for more space
        using (FileStream fs = new FileStream(fullPath, FileMode.Create))
        {
            iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
            doc.Open();

            // --- Header & Logo ---
            System.Drawing.Image resImage = Properties.Resources.logo1;
            iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(resImage, System.Drawing.Imaging.ImageFormat.Png);
            logo.ScaleToFit(70f, 70f);
            logo.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
            doc.Add(logo);

            var titleFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 16);
            var subFont = iTextSharp.text.FontFactory.GetFont("Helvetica-Bold", 11);
            var regFont = iTextSharp.text.FontFactory.GetFont("Helvetica", 9);

            iTextSharp.text.Paragraph h1 = new iTextSharp.text.Paragraph("MARONDERA PROVINCIAL HOSPITAL", titleFont);
            h1.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
            doc.Add(h1);

            iTextSharp.text.Paragraph h2 = new iTextSharp.text.Paragraph("DETAILED PATIENT MOVEMENT & BED STATISTICS REPORT", subFont);
            h2.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
            doc.Add(h2);
            doc.Add(new iTextSharp.text.Paragraph("Reporting Period: " + start.ToString("dd/MM/yyyy") + " to " + end.ToString("dd/MM/yyyy") + " | Generated By: " + currentUser, regFont));
            doc.Add(new iTextSharp.text.Paragraph("\n"));

            // --- SECTION 1: STATISTICAL SUMMARY ---
            doc.Add(new iTextSharp.text.Paragraph("1. EXECUTIVE SUMMARY", subFont));
            doc.Add(new iTextSharp.text.Paragraph("-" + new string('-', 120), regFont));
            
            // Calculate totals from the DataTable locally
            int adm = movementData.Select("category = 'Admission'").Length;
            int dis = movementData.Select("category = 'Discharge'").Length;
            int trans = movementData.Select("category LIKE 'InterWard%'").Length;
            int death = movementData.Select("category = 'Death'").Length;

            doc.Add(new iTextSharp.text.Paragraph("Total Admissions: " + adm, regFont));
            doc.Add(new iTextSharp.text.Paragraph("Total Discharges: " + dis, regFont));
            doc.Add(new iTextSharp.text.Paragraph("Total Transfers: " + trans, regFont));
            doc.Add(new iTextSharp.text.Paragraph("Total Deaths: " + death, regFont));
            doc.Add(new iTextSharp.text.Paragraph("\n"));

            // --- SECTION 2: DETAILED LOG ---
            doc.Add(new iTextSharp.text.Paragraph("2. DETAILED MOVEMENT LOG", subFont));
            doc.Add(new iTextSharp.text.Paragraph("\n"));

            iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(7);
            table.WidthPercentage = 100;
            table.SetWidths(new float[] { 2f, 1.5f, 2f, 2f, 1.5f, 1.5f, 1.5f });

            string[] headers = { "Date/Time", "Hosp #", "Name", "Surname", "Category", "From", "To" };
            foreach (string h in headers)
            {
                iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(h, subFont));
                cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }

            foreach (DataRow row in movementData.Rows)
            {
                table.AddCell(new iTextSharp.text.Phrase(Convert.ToDateTime(row["MovementDateTime"]).ToString("g"), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["hospitalNumber"].ToString(), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["name"].ToString(), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["surname"].ToString(), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["category"].ToString(), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["fromWard"].ToString(), regFont));
                table.AddCell(new iTextSharp.text.Phrase(row["toWard"].ToString(), regFont));
            }
            doc.Add(table);

            // --- Signature ---
            doc.Add(new iTextSharp.text.Paragraph("\n\n"));
            iTextSharp.text.Paragraph sig = new iTextSharp.text.Paragraph("___________________________\nRecords Officer Signature", regFont);
            sig.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            doc.Add(sig);

            doc.Close();
        }

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullPath) { UseShellExecute = true });
    }
    catch (Exception ex)
    {
        MessageBox.Show("Detailed Export Error: " + ex.Message);
    }

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
       