using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Data.OleDb;
using Microsoft.Reporting.WinForms;

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
            comboBox2.Items.Add("Female Ward");
            comboBox2.Items.Add("Labor Ward");
            comboBox2.Items.Add("Male Ward");
            comboBox2.Items.Add("DECEASED");
          
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string criteria = comboBox1.SelectedItem.ToString();
                string value = textBox1.Text.Trim();

                if (string.IsNullOrWhiteSpace(value))
                {
                    MessageBox.Show("Please enter a value to search for.", "Warning",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string query = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                               "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                               "WHERE ";
                switch (criteria)
                {
                    case "Hospital Number":
                        query += "PM.[hospitalNumber] = ?";
                        break;
                    case "Name":
                        query += "PM.[name] LIKE ?";
                        value = "%" + value + "%";
                        break;
                    case "Surname":
                        query += "PM.[surname] LIKE ?";
                        value = "%" + value + "%";
                        break;
                    case "Category":
                        query += "TM.[category] LIKE ?";
                        value = "%" + value + "%";
                        break;
                    case "Ward":
                        query += "PM.[currentWard] LIKE ?";
                        value = "%" + value + "%";
                        break;
                    default:
                        MessageBox.Show("Invalid search criteria selected: " + criteria, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                }


                  using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                    {
                        con.Open();

                        using (OleDbCommand cmd = new OleDbCommand(query, con))
                        {
                        }

                        con.Close();
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
               
            }
        
       
        private void button5_Click(object sender, EventArgs e)
        {
           
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

    if (!DateTime.TryParse(textBox1.Text, out startDate) ||
        !DateTime.TryParse(textBox2.Text, out endDate))
    {
        button1.Enabled = false;
        return;
    }

    button1.Enabled = true;

    if (comboBox2.SelectedItem == null)
        return;

    string ward = comboBox2.SelectedItem.ToString();

    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        try
        {
            con.Open();

            // ===============================
            // CURRENT ADMITTED COUNT
            // ===============================
            using (OleDbCommand cmd = new OleDbCommand(
                "SELECT COUNT(*) FROM tblPatientMaster WHERE currentWard = ? AND isAdmitted = 'Yes'", con))
            {
                cmd.Parameters.AddWithValue("?", ward);
                label5.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
            }

            // ===============================
            // CURRENT DISCHARGED COUNT
            // ===============================
            using (OleDbCommand cmd = new OleDbCommand(
                "SELECT COUNT(*) FROM tblPatientMaster WHERE currentWard = ? AND isAdmitted = 'No'", con))
            {
                cmd.Parameters.AddWithValue("?", ward);
                label7.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
            }

            // ===============================
            // TOTAL IN (toWard) — ACCESS SAFE
            // ===============================
            using (OleDbCommand cmd = new OleDbCommand(
                @"SELECT COUNT(*) FROM (
                      SELECT DISTINCT hospitalNumber
                      FROM tblPatientMovement
                      WHERE toWard = ?
                        AND MovementDateTime BETWEEN ? AND ?
                  )", con))
            {
                cmd.Parameters.AddWithValue("?", ward);
                cmd.Parameters.AddWithValue("?", startDate);
                cmd.Parameters.AddWithValue("?", endDate);
                label6.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
            }

            // ===============================
            // TOTAL OUT (fromWard) — ACCESS SAFE
            // ===============================
            using (OleDbCommand cmd = new OleDbCommand(
                @"SELECT COUNT(*) FROM (
                      SELECT DISTINCT hospitalNumber
                      FROM tblPatientMovement
                      WHERE fromWard = ?
                        AND MovementDateTime BETWEEN ? AND ?
                  )", con))
            {
                cmd.Parameters.AddWithValue("?", ward);
                cmd.Parameters.AddWithValue("?", startDate);
                cmd.Parameters.AddWithValue("?", endDate);
                label8.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
            }

            // ===============================
            // CURRENT DEATH COUNT
            // ===============================
            using (OleDbCommand cmd = new OleDbCommand(
                "SELECT COUNT(*) FROM tblPatientMaster WHERE currentWard = ? AND isAdmitted = 'No'", con))
            {
                cmd.Parameters.AddWithValue("?", ward);
                label9.Text = Convert.ToInt32(cmd.ExecuteScalar()).ToString();
            }

        }
        catch (Exception error)
        {
            MessageBox.Show(
                "Error retrieving statistics:\n" + error.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

}

        

        private void button11_Click(object sender, EventArgs e)
        {
            textBox16.Clear();
            
        }

        private void button12_Click(object sender, EventArgs e)
        {
           try{
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
                    con.Close();
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
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    string sql = @"SELECT [name], [surname], [password]
               FROM userAccounts
               WHERE [idNumber]=?";

                    OleDbCommand com = new OleDbCommand(sql, con);
                    com.Parameters.AddWithValue("?", textBox19.Text); // idNumber

                    OleDbDataReader dr = com.ExecuteReader();

                    if (dr.Read())
                    {
                        textBox17.Text = dr["name"].ToString();
                        textBox18.Text = dr["surname"].ToString();
                        textBox20.Text = dr["password"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("Error retrieving statistics:\n", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    dr.Close();
                    con.Close();
                }
            }
               
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception error)
            {
                MessageBox.Show("Error retrieving statistics:\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
}