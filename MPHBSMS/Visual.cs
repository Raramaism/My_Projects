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
using System.Windows.Forms.DataVisualization.Charting;

namespace MPHBSMS
{
    public partial class Visual : Form
    {

        public Visual()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 obj = new Form1();
            this.Close();
            obj.Show();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Menu obj = new Menu();
            this.Close();
            obj.Show();
           

        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                con.Open();
                string com = "select * from tblPatientMaster";
                OleDbCommand comm = new OleDbCommand(com, con);
                OleDbDataAdapter da = new OleDbDataAdapter(comm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.Refresh();
                dataGridView2.Visible = false;
                dataGridView1.Visible = true;
                chart1.Visible = false;
                con.Close();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void mentalHealthUnitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Define a map for colors aligned with categories (for the chart)
            var categoryColors = new Dictionary<string, System.Drawing.Color>
{
    {"Admission", System.Drawing.Color.MediumSeaGreen},
    {"InterWardTransferIn", System.Drawing.Color.SteelBlue},
    {"Discharge", System.Drawing.Color.Firebrick},
    {"InterWardTransferOut", System.Drawing.Color.Orange},
    {"Death", System.Drawing.Color.DarkRed}
};

            // Initializing visibility and clearing controls
            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            textBox1.Clear();


            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (Using Indexing for Color)
                    // ----------------------------------------------------

                    // Ensure chart is clear before adding new data
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;

                    // EFFICIENT QUERY: Get all counts in one database call
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE PM.currentWard = ? " ;

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        com.Parameters.AddWithValue("?", ward);

                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            // We'll track the index of the point as we add it
                            int pointIndex = 0;

                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                // Add point (no local 'point' variable needed)
                                activitySeries.Points.AddXY(category, count);

                                // Get a direct reference to the LAST ADDED point via its index
                                System.Drawing.Color barColor = System.Drawing.Color.Gray;

                                if (categoryColors.ContainsKey(category))
                                {
                                    barColor = categoryColors[category];
                                }

                                // Set the color property directly on the point at the current index
                                activitySeries.Points[pointIndex].Color = barColor;

                                pointIndex++;
                            }
                        }
                    }

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (CORRECTED QUERY)
                    // ----------------------------------------------------

                    // Query to display all joined patient and movement details for the selected ward
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = ?";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void femaleWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWardTransferIn", "InterWardTransferIn"},
    {"Discharge", "Discharge"},
    {"InterWardTransferOut", "InterWardTransferOut"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber WHERE PM.currentWard = ? AND TM.category = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void paediatricWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMovement WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void maleWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMovement WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            } 
        }

        private void postNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMovement WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void neoNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMovement WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void antiNatalWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMovement WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMovement  WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void laborWardToolStripMenuItem_Click(object sender, EventArgs e)
        {

            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMaster WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitalNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMaster WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void accidentAndEmergencyToolStripMenuItem_Click(object sender, EventArgs e)
        {


            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            var metrics = new Dictionary<string, string>
{
    {"Admissions", "Admission"},
    {"InterWard Transfer In", "Inter Ward Transfer In"},
    {"Discharge", "Discharge"},
    {"Inter Ward Transfer Out", "Inter Ward Transfer Out"},
    {"Death", "Death"}
};

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();
                    chart1.Series.Clear();

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    var activitySeries = chart1.Series.Add("Activity Count");
                    activitySeries.ChartType = SeriesChartType.Column;
                    activitySeries.IsValueShownAsLabel = true;
                    activitySeries.XValueType = ChartValueType.String;

                    foreach (var metric in metrics)
                    {
                        string seriesName = metric.Key;
                        string categoryValue = metric.Value;

                        string commandText = "SELECT COUNT(*) FROM tblPatientMaster WHERE [ward] = ? AND [category] = ?";
                        using (OleDbCommand com = new OleDbCommand(commandText, con))
                        {
                            com.Parameters.AddWithValue("?", ward);
                            com.Parameters.AddWithValue("?", categoryValue);

                            int count = 0;
                            object result = com.ExecuteScalar();

                            if (result != null && result != DBNull.Value)
                            {
                                count = Convert.ToInt32(result);
                            }


                            activitySeries.Points.AddXY(seriesName, count);
                        }
                    }

                    string myquaery = "SELECT [hospitlNumber],[name],[surname],[gender],[date],[time],[category] FROM tblPatientMaster WHERE [ward] = ?"; // Selective columns recommended
                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    comm.Parameters.AddWithValue("?", ward);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void Visual_Load(object sender, EventArgs e)
        {
            DatabaseHelper.InitializeDatabase();
            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                con.Open();
               // string com = "SELECT tblPatientMaster.hospitalNumber,tblPatientMaster.name,tblPatientMaster.surname,tblPatientMaster.gender,tblPatientMaster.currentWard, tblPatientMaster.isAdmitted,tblPatientMaster.admissionDate,tblPatientMaster.dischargeDate,tblPatientMovement.MovementDateTime,tblPatientMovement.toWard,tblPatientMovement.fromWard,tblPatientMovement.category,tblPatientMovement.enteredBy FROM tblPatientMaster JOIN tblPatientMovement ON tblPatientMaster.hospitalNumber = tblPatientMovement.hospitalNumber ORDER BY tblPatientMaster.hospitalNumber ASC"; 
              string com = " SELECT tblPatientMaster.hospitalNumber, tblPatientMaster.name, tblPatientMaster.surname, tblPatientMaster.gender, tblPatientMaster.currentWard, tblPatientMaster.isAdmitted, tblPatientMaster.admissionDate, tblPatientMaster.dischargeDate, tblPatientMovement.MovementDateTime, tblPatientMovement.toWard, tblPatientMovement.fromWard, tblPatientMovement.category, tblPatientMovement.enteredBy FROM tblPatientMaster INNER JOIN tblPatientMovement ON tblPatientMaster.hospitalNumber = tblPatientMovement.hospitalNumber ORDER BY tblPatientMaster.hospitalNumber ASC";
                OleDbCommand comm = new OleDbCommand(com, con);
                OleDbDataAdapter da = new OleDbDataAdapter(comm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                chart1.Visible = false;
                dataGridView2.Visible = false;

                comboBox1.Items.AddRange(new string[]
    {
        "Hospital Number",
        "Name",
        "Surname",
        "Category",
        "Ward"
    });
                comboBox1.SelectedIndex = 0; // Default selection

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.ReadOnly = true;
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                con.Close();
            }

            using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
            {
                con.Open();
                string com = "select * from tblPatientMaster";
                OleDbCommand comm = new OleDbCommand(com, con);
                OleDbDataAdapter da = new OleDbDataAdapter(comm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                chart1.Visible = false;
                dataGridView2.Visible = false;

                comboBox1.Items.AddRange(new string[]
    {
        "Hospital Number",
        "Name",
        "Surname",
        "Category",
        "Ward"
    });
                comboBox1.SelectedIndex = 0; // Default selection

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.ReadOnly = true;
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                con.Close();
            }

        }

        private void chart1_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
    string criteria = comboBox1.SelectedItem.ToString();
    string value = textBox1.Text.Trim();

    if (string.IsNullOrWhiteSpace(value))
    {
        MessageBox.Show("Please enter a value to search for.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }

    string query = "SELECT * FROM tblPatientMaster  WHERE ";

    switch (criteria)
    {
        case "Hospital Number":
            query += "[hospitalNumber] = ?";
            break;
        case "Name":
            query += "[name] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Surname":
            query += "[surname] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Category":
            query += "[category] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Ward":
            query += "[ward] LIKE ?";
            value = "%" + value + "%";
            break;
        default: 
            MessageBox.Show("Invalid search criteria selected: " + criteria, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return; 
    }

    try
    {
        using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
        {
            con.Open();

            using (OleDbCommand cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("?", value);

                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    dataGridView2.DataSource = dt;
                    chart1.Visible = false;
                    dataGridView1.Visible = false;
                    dataGridView2.Visible = true;
                }
                else
                {
                    MessageBox.Show("No matching records found.", "Search Results",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dataGridView2.DataSource = null;
                }
            }

            con.Close();
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error performing search:\n" + ex.Message, "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
