namespace MPHBSMS
{
    partial class Reports
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mentalHealthUnitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.femaleWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.paedToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.maleWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.postNatalWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.neoNatalWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.antiNatalWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.laborWardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.accidentAndEmeregencyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Location = new System.Drawing.Point(163, 46);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(396, 246);
            this.reportViewer1.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mentalHealthUnitToolStripMenuItem,
            this.femaleWardToolStripMenuItem,
            this.paedToolStripMenuItem,
            this.maleWardToolStripMenuItem,
            this.postNatalWardToolStripMenuItem,
            this.neoNatalWardToolStripMenuItem,
            this.antiNatalWardToolStripMenuItem,
            this.laborWardToolStripMenuItem,
            this.accidentAndEmeregencyToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(870, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mentalHealthUnitToolStripMenuItem
            // 
            this.mentalHealthUnitToolStripMenuItem.Name = "mentalHealthUnitToolStripMenuItem";
            this.mentalHealthUnitToolStripMenuItem.Size = new System.Drawing.Size(119, 20);
            this.mentalHealthUnitToolStripMenuItem.Text = "Mental Health Unit";
            // 
            // femaleWardToolStripMenuItem
            // 
            this.femaleWardToolStripMenuItem.Name = "femaleWardToolStripMenuItem";
            this.femaleWardToolStripMenuItem.Size = new System.Drawing.Size(88, 20);
            this.femaleWardToolStripMenuItem.Text = "Female Ward";
            // 
            // paedToolStripMenuItem
            // 
            this.paedToolStripMenuItem.Name = "paedToolStripMenuItem";
            this.paedToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.paedToolStripMenuItem.Text = "Paed";
            // 
            // maleWardToolStripMenuItem
            // 
            this.maleWardToolStripMenuItem.Name = "maleWardToolStripMenuItem";
            this.maleWardToolStripMenuItem.Size = new System.Drawing.Size(76, 20);
            this.maleWardToolStripMenuItem.Text = "Male Ward";
            // 
            // postNatalWardToolStripMenuItem
            // 
            this.postNatalWardToolStripMenuItem.Name = "postNatalWardToolStripMenuItem";
            this.postNatalWardToolStripMenuItem.Size = new System.Drawing.Size(104, 20);
            this.postNatalWardToolStripMenuItem.Text = "Post Natal Ward";
            // 
            // neoNatalWardToolStripMenuItem
            // 
            this.neoNatalWardToolStripMenuItem.Name = "neoNatalWardToolStripMenuItem";
            this.neoNatalWardToolStripMenuItem.Size = new System.Drawing.Size(100, 20);
            this.neoNatalWardToolStripMenuItem.Text = "NeoNatal Ward";
            // 
            // antiNatalWardToolStripMenuItem
            // 
            this.antiNatalWardToolStripMenuItem.Name = "antiNatalWardToolStripMenuItem";
            this.antiNatalWardToolStripMenuItem.Size = new System.Drawing.Size(100, 20);
            this.antiNatalWardToolStripMenuItem.Text = "AntiNatal Ward";
            // 
            // laborWardToolStripMenuItem
            // 
            this.laborWardToolStripMenuItem.Name = "laborWardToolStripMenuItem";
            this.laborWardToolStripMenuItem.Size = new System.Drawing.Size(80, 20);
            this.laborWardToolStripMenuItem.Text = "Labor Ward";
            this.laborWardToolStripMenuItem.Click += new System.EventHandler(this.laborWardToolStripMenuItem_Click);
            // 
            // accidentAndEmeregencyToolStripMenuItem
            // 
            this.accidentAndEmeregencyToolStripMenuItem.Name = "accidentAndEmeregencyToolStripMenuItem";
            this.accidentAndEmeregencyToolStripMenuItem.Size = new System.Drawing.Size(157, 20);
            this.accidentAndEmeregencyToolStripMenuItem.Text = "Accident and Emeregency";
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // Reports
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(870, 372);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Reports";
            this.Text = "Marondera Provincial Hospital Bed Statistics Management System ";
            this.Load += new System.EventHandler(this.Reports_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mentalHealthUnitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem femaleWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem paedToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem maleWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem postNatalWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem neoNatalWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem antiNatalWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem laborWardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem accidentAndEmeregencyToolStripMenuItem;
        private System.Windows.Forms.PrintDialog printDialog1;
    }
}