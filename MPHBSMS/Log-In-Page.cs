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
    public partial class Log_In_Page : Form
    {
        public Log_In_Page()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MPHBSMS.accdb");
                con.Open();

                string logIn = "select[name],[identityNumber],[password] FROM accounts WHERE ([identityNumber] = ? AND [password]=?)";
                OleDbCommand com = new OleDbCommand(logIn, con);

                com.Parameters.AddWithValue("@identyNumber", textBox1.Text);
                com.Parameters.AddWithValue("@password", textBox2.Text);
                OleDbDataReader reader = com.ExecuteReader();

                if (reader.Read() == true)
                {

                    Menu menupage = new Menu();
                    this.Hide();
                    menupage.ShowDialog();
                    this.Show();

                    string name = reader[0].ToString();

                    DialogResult = MessageBox.Show("Welcome " + name + "!", "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                }
                else
                {
                    Menu menupage = new Menu();
                    this.Hide();
                    menupage.ShowDialog();
                    this.Show();

                    textBox1.Clear();
                    textBox2.Clear();
                    DialogResult = MessageBox.Show("User not found!", "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                }

                con.Close();
            }

            catch (Exception error)
            {
                DialogResult = MessageBox.Show("An error occured\n" + error.Message, "Marondera Provincial Hospital", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.PasswordChar='*';
        }

        private void Log_In_Page_Load(object sender, EventArgs e)
        {

        }
    }
}
