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
using System.IO;

namespace MPHBSMS
{
    public partial class Log_In_Page : Form
    {
        public Log_In_Page()
        {
            InitializeComponent();
            DatabaseHelper.InitializeDatabase();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    textBox1.Focus();
                    string logIn = "select [name],[surname],[idNumber],[password] FROM userAccounts WHERE ([idNumber] = ? AND [password]=?)";
                    OleDbCommand com = new OleDbCommand(logIn, con);

                    com.Parameters.AddWithValue("@idNumber", textBox1.Text.Trim());
                    com.Parameters.AddWithValue("@password", textBox2.Text.Trim());
                    OleDbDataReader reader = com.ExecuteReader();

                    if (reader.Read() == true)
                    {
                        string name = reader[0].ToString();
                        string surname = reader[1].ToString();

                        
                      MessageBox.Show("Welcome " + name + " " + surname + "!", "Marondera Provincial Hospital", MessageBoxButtons.OK, MessageBoxIcon.Information);
                       
                        MPHBSMS.CurrentUser = reader[0].ToString().Trim();

                        Menu menupage = new Menu();
                        menupage.Show();
                        this.Hide(); 
                    } else 
                    {
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox1.Focus();
                         MessageBox.Show("User not found!", "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                       
                    }

                    con.Close();
                }
                
            }

            catch (Exception error)
            {
                 MessageBox.Show("An error occured\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox1.Focus();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Log_In_Page_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            textBox2.PasswordChar = '*';
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
