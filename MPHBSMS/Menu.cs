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
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MPHBSMS.accdb");
        

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
           
        }

        private void Menu_Load(object sender, EventArgs e)
        {
            try
            {
                textBox1.Focus();
              
                con.Open();

                    string query1 = ("Select Count(*)From mphtable");
                    OleDbCommand com1 = new OleDbCommand(query1, con);
                    int broughtForward = (int)com1.ExecuteScalar();
                    label15.Text = broughtForward.ToString();

                    string query7 = ("Select Count(*)From mphtable");
                    OleDbCommand com7 = new OleDbCommand(query7, con);
                    int bedsOccupied = (int)com7.ExecuteScalar();
                    label21.Text = bedsOccupied.ToString();
            }
            catch(Exception error){
               
                DialogResult = MessageBox.Show("ERROR!!\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
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
                con.Open();
                string gender = "";
                string category = "";
                string ward = "";
                const string CatAdmission = "Admission";
                const string CatTransferIn = "Inter Transfer In";
                const string CatDischarge = "Discharge";
                const string CatTransferOut = "Inter Transfer Out";
                const string CatDeath = "Death";

               if(radioButton1.Checked)
               {
                   gender = "female";
                   
               }
               else if (radioButton2.Checked)
               {
                   gender = "male";
                  
               }
                if(checkBox1.Checked)
               {
                   category = "Admission";
               }else if(checkBox2.Checked)
               {
                   category = "Inter Transfer In";
               }else if(checkBox3.Checked)
               {
                   category = "Discharge";
               }else if(checkBox4.Checked)
               {
                   category = "Inter Transfer Out";
               }else if(checkBox5.Checked)
               {
                   category = "Death";
               }
               
                ward = listBox1.SelectedItem.ToString();

               string date = dateTimePicker1.Value.ToShortDateString();
               string time = dateTimePicker1.Value.ToShortTimeString();
                string currentDate = DateTime.Now.ToShortDateString();

                string query2 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] = ?";
                    OleDbCommand com2 = new OleDbCommand(query2, con);
                    com2.Parameters.AddWithValue("?", currentDate);
                    com2.Parameters.AddWithValue("?", CatAdmission);
                    int admissions = (int)com2.ExecuteScalar();
                    label16.Text = admissions.ToString();

                    string query3 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] = ?";
                    OleDbCommand com3 = new OleDbCommand(query3, con);
                    com3.Parameters.AddWithValue("?", currentDate);
                    com3.Parameters.AddWithValue("?", CatTransferIn);
                    int interWardsTransferIn = (int)com3.ExecuteScalar();
                    label17.Text = interWardsTransferIn.ToString();

                    string query4 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] = ?";
                    OleDbCommand com4 = new OleDbCommand(query4, con);
                    com4.Parameters.AddWithValue("?", currentDate);
                    com4.Parameters.AddWithValue("?", CatDischarge);
                    int discharge = (int)com4.ExecuteScalar();
                    label18.Text = discharge.ToString();

                    string query5 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] = ?";
                    OleDbCommand com5 = new OleDbCommand(query5, con);
                    com5.Parameters.AddWithValue("?", currentDate);
                    com5.Parameters.AddWithValue("?", CatTransferOut);
                    int interWardsTransferOut = (int)com5.ExecuteScalar();
                    label19.Text = interWardsTransferOut.ToString();

                    string query6 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] = ?";
                    OleDbCommand com6 = new OleDbCommand(query6, con);
                    com6.Parameters.AddWithValue("?", currentDate);
                    com6.Parameters.AddWithValue("?", CatDeath);
                    int deaths = (int)com6.ExecuteScalar();
                    label20.Text = deaths.ToString();

                 
                    string previousDate = DateTime.Now.AddDays(-1).ToShortDateString();
                    string query1 = "SELECT Count(*) FROM mphtable WHERE [date] = ?";
                    OleDbCommand com1 = new OleDbCommand(query1, con);
                    com1.Parameters.AddWithValue("?", previousDate);
                    int broughtForward = (int)com1.ExecuteScalar();
                    label15.Text = broughtForward.ToString();

                    const string CatDischarges = "Discharge";
                    const string CatDeaths = "Death";
                    string query7 = "SELECT Count(*) FROM mphtable WHERE [date] = ? AND [category] NOT IN (?, ?)";
                    OleDbCommand com7 = new OleDbCommand(query7, con);
                    com7.Parameters.AddWithValue("?", previousDate);
                    com7.Parameters.AddWithValue("?", CatDeaths);
                    com7.Parameters.AddWithValue("?", CatDischarges);
                    int bedsOccupied = (int)com7.ExecuteScalar();
                    label21.Text = bedsOccupied.ToString();

                OleDbCommand com = new OleDbCommand("insert into mphtable ([hospitalNumber],[name],[surname],[gender],[to],[from],[date],[time],[category],[ward]) values (?,?,?,?,?,?,?,?,?,?)", con);
                com.Parameters.AddWithValue("@hospitalNumber", textBox1.Text);
                com.Parameters.AddWithValue("@name", textBox2.Text);
                com.Parameters.AddWithValue("@surname", textBox3.Text);
                com.Parameters.AddWithValue("@gender",gender);
                com.Parameters.AddWithValue("@to",ward);
                com.Parameters.AddWithValue("@from",textBox4.Text);
                com.Parameters.AddWithValue("@date",date);
                com.Parameters.AddWithValue("@time", time);
                com.Parameters.AddWithValue("@category",category);
                com.Parameters.AddWithValue("@ward", selectedWard);

                
                if (string.IsNullOrWhiteSpace(textBox1.Text) && string.IsNullOrWhiteSpace(textBox2.Text) && string.IsNullOrWhiteSpace(textBox3.Text) && string.IsNullOrWhiteSpace(textBox4.Text))
                {
                    DialogResult = MessageBox.Show("Please fill in the textboxes\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
              
                else if (!radioButton1.Checked && !radioButton2.Checked)
                {
                    DialogResult = MessageBox.Show("Please select GENDER\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
               
                else if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked)
                {
                    
                    DialogResult = MessageBox.Show("Please select one CATEGORY\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!checkBox6.Checked)
                {
                    DialogResult = MessageBox.Show("Complete your action!!\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (string.IsNullOrEmpty(ward))
                {
                    DialogResult = MessageBox.Show("Please select a Ward\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    int rowsAffected = com.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        DialogResult = MessageBox.Show("Record inserted successfully\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                
            }
    catch (Exception error)
            {
        DialogResult = MessageBox.Show("ERROR!!\n"+error.Message,"Marondera Provincial Hospital",MessageBoxButtons.RetryCancel,MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }

        }

        private void mentalHealthUnitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            selectedWard = ClickedItem.Text;
            DialogResult = MessageBox.Show("You selected: \n" + selectedWard , "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
