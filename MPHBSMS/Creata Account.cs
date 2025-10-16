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
    public partial class Creata_Account : Form
    {
        public Creata_Account()
        {
            InitializeComponent();

            textBox6.Visible = false;
            label7.Visible = false;
            textBox1.Focus();

           
            textBox1.TextChanged += CheckFields;
            textBox2.TextChanged += CheckFields;
            textBox3.TextChanged += CheckFields;
            textBox4.TextChanged += CheckFields;
            textBox5.TextChanged += CheckFields;
        }

        private void CheckFields(object sender, EventArgs e)
        {
            
            bool allFilled = !string.IsNullOrWhiteSpace(textBox1.Text) &&
                             !string.IsNullOrWhiteSpace(textBox2.Text) &&
                             !string.IsNullOrWhiteSpace(textBox3.Text) &&
                             !string.IsNullOrWhiteSpace(textBox4.Text) &&
                             !string.IsNullOrWhiteSpace(textBox5.Text);

           
            if (allFilled)
            {
                label7.Visible = true;
                textBox6.Visible = true;
            }
            else
            {
                label7.Visible = false;
                textBox6.Visible = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Form1 Obj = new Form1();
                
                this.Close();
                Obj.Show();
            }
                  
    catch (Exception error)
{
   
        DialogResult = MessageBox.Show("ERROR!!\n"+error.Message,"Marondera Provincial Hospital",MessageBoxButtons.RetryCancel,MessageBoxIcon.Error);
}

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Creata_Account_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
            }
                    
    catch (Exception error)
{
   
        DialogResult = MessageBox.Show("ERROR!!\n"+error.Message,"Marondera Provincial Hospital",MessageBoxButtons.RetryCancel,MessageBoxIcon.Error);
}

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DatabaseHelper.InitializeDatabase();
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    string account = "insert into accounts ([name],[surname],[identityNumber],[password]) values (?,?,?,?)";
                    OleDbCommand com = new OleDbCommand(account, con);
                    com.Parameters.AddWithValue("?", textBox1.Text);
                    com.Parameters.AddWithValue("?", textBox2.Text);
                    com.Parameters.AddWithValue("?", textBox3.Text);
                    com.Parameters.AddWithValue("?", textBox5.Text);
                    if (textBox6.Text == "MPHb0x20M0H&CC")
                    {
                        com.ExecuteNonQuery();
                        DialogResult = MessageBox.Show("Admin password is incorrect\n", "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    }
                    else
                    {
                        com.ExecuteNonQuery();
                        DialogResult = MessageBox.Show("Account created successfully\n", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox3.Clear();
                        textBox4.Clear();
                        textBox5.Clear();
                        textBox6.Clear();
                        textBox1.Focus();


                        Log_In_Page bj = new Log_In_Page();
                        this.Hide();
                        bj.ShowDialog();
                        this.Show();
                    }

                    con.Close();

                   
                }
            }

            catch (Exception error)
            {

                DialogResult = MessageBox.Show("ERROR!!\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            textBox6.PasswordChar = '*';
        }

        private void label7_Click(object sender, EventArgs e)
        {
            
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            textBox5.PasswordChar = '*';
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox4.PasswordChar = '*';
        }
    }
}
