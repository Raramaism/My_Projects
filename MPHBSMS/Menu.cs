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

            try
{
    using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
    {
        con.Open();

        string gender = "";
        string category = "";
        string ward = "";
        string home = "";

        const string CatAdmission = "Admission";
        const string CatTransferIn = "Inter Ward Transfer In";
        const string CatDischarge = "Discharge";
        const string CatTransferOut = "Inter Ward Transfer Out";
        const string CatDeath = "Death";

        if (radioButton1.Checked)
            gender = "female";
        else if (radioButton2.Checked)
            gender = "male";

        if (checkBox1.Checked)
            category = CatAdmission;
        else if (checkBox2.Checked)
            category = CatTransferIn;
        else if (checkBox3.Checked)
            category = CatDischarge;
        else if (checkBox4.Checked)
            category = CatTransferOut;
        else if (checkBox5.Checked)
            category = CatDeath;

        ward = listBox1.SelectedItem.ToString();
        home = listBox2.SelectedItem.ToString();

        string date = dateTimePicker1.Value.ToShortDateString();
        string time = dateTimePicker1.Value.ToString("h:mmtt").ToLower();
        string hospitalNumber = textBox1.Text.Trim();

        if (string.IsNullOrWhiteSpace(hospitalNumber) ||
            string.IsNullOrWhiteSpace(textBox2.Text) ||
            string.IsNullOrWhiteSpace(textBox3.Text))
        {
            MessageBox.Show("Please fill in all required textboxes.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrEmpty(gender))
        {
            MessageBox.Show("Please select GENDER.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrEmpty(category))
        {
            MessageBox.Show("Please select a CATEGORY.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!checkBox6.Checked)
        {
            MessageBox.Show("Complete your action!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrEmpty(ward))
        {
            MessageBox.Show("Please select a Ward.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        bool patientExists = false;
        string checkQuery = "SELECT COUNT(*) FROM mphtable WHERE [hospitalNumber] = ?";
        using (OleDbCommand checkCmd = new OleDbCommand(checkQuery, con))
        {
            checkCmd.Parameters.AddWithValue("?", hospitalNumber);
            int count = (int)checkCmd.ExecuteScalar();
            patientExists = count > 0;
        }

        if (category == CatAdmission)
        {
            if (patientExists)
            {
                MessageBox.Show("This patient is already admitted.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        else
        {
            if (!patientExists)
            {
                MessageBox.Show("Patient does not exist in the database.", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        string insertQuery = "INSERT INTO mphtable ([hospitalNumber],[name],[surname],[gender],[to],[from],[date],[time],[category],[ward],[enteredBy]) VALUES (?,?,?,?,?,?,?,?,?,?,?)";
      
        using (OleDbCommand com = new OleDbCommand(insertQuery, con))
{
    com.Parameters.AddWithValue("@hospitalNumber", hospitalNumber);
    com.Parameters.AddWithValue("@name", textBox2.Text);
    com.Parameters.AddWithValue("@surname", textBox3.Text);
    com.Parameters.AddWithValue("@gender", gender);
    com.Parameters.AddWithValue("@to", ward);
    com.Parameters.AddWithValue("@from", home);
    com.Parameters.AddWithValue("@date", date);
    com.Parameters.AddWithValue("@time", time);
    com.Parameters.AddWithValue("@category", category);
    com.Parameters.AddWithValue("@ward", ward);
    com.Parameters.AddWithValue("@enteredBy", MPHBSMS.CurrentUser);

    int rows = com.ExecuteNonQuery();

    if (rows > 0)
    {
        MessageBox.Show("Record inserted successfully by "+MPHBSMS.CurrentUser,
                        "Marondera Provincial Hospital",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
    }
}
        con.Close();
        }
    }

catch (Exception error)
{
    MessageBox.Show("ERROR!!\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
}

        }    

        private void mentalHealthUnitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard, "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
          
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
    }
}
