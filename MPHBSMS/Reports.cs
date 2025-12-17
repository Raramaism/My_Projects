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

          
        }

        private void button4_Click(object sender, EventArgs e)
        {

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

        }

        private void button5_Click_2(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            textBox16.Clear();
            
        }

        private void button12_Click(object sender, EventArgs e)
        {
            
        }
    }
}