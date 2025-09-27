using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPHBSMS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception error)
            {
                DialogResult = MessageBox.Show("An Error Occured. See Details Below:\n" + error.Message, "Health Information Management System" ,MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Log_In_Page Obj = new Log_In_Page();
                this.Hide();
                Obj.ShowDialog();
                this.Show();
            }
            catch (Exception error)
            {
                DialogResult = MessageBox.Show("An Error Occured. See Details Below:\n" + error.Message, "Health Information Management System", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Creata_Account Obj = new Creata_Account();
                this.Hide();
                Obj.ShowDialog();
                this.Show();
            }
            catch (Exception error)
            {
                DialogResult = MessageBox.Show("An Error Occured. See Details Below:\n" + error.Message, "Health Information Management System", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception error)
            {
                DialogResult = MessageBox.Show("An Error Occured. See Details Below:\n" + error.Message, "Health Information Management System", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Menu menupage = new Menu();
            this.Hide();
            menupage.ShowDialog();
            this.Show();
        }
    }
}
