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
                Log_In_Page bj = new Log_In_Page();
                this.Hide();
                bj.ShowDialog();
                
                
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

                Creata_Account obj = new Creata_Account();
                this.Show();
                obj.ShowDialog();
                this.Hide();
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
            this.Close();
            menupage.Show();
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            BackupManager.ExecuteAdvancedBackup();
            // This captures the logout even if they don't click the logout button
            if (MPHBSMS.CurrentUser != "")
            {
                MPHBSMS.LogLogout();
            }
        }

        private void backupTimer_Tick(object sender, EventArgs e)
        {
           

            // Check for Shift Change at 4:00 PM
            if (DateTime.Now.Hour == 16 && DateTime.Now.Minute == 0)
            {
                // Prevent multiple backups in the same minute
                backupTimer.Enabled = false;

                BackupManager.ExecuteAdvancedBackup();


                // Re-enable after the minute passes
                System.Threading.Thread.Sleep(61000);
                backupTimer.Enabled = true;
            }
        }
}
}

