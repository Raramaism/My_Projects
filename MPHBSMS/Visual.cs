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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
            // Assuming textBox1 is a valid control
            // textBox1.Clear(); 

            ToolStripMenuItem ClickedItem = (ToolStripMenuItem)sender;
            var ward = ClickedItem.Text;

            // Sanitize the ward variable to prevent SQL errors
            string safeWard = ward.Replace("'", "''");

            try
            {
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND DATA LOADING (FIXED FOR LEGEND AND GRIDLINES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    // Ensure gridlines are enabled (default behavior)
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // EFFICIENT QUERY to get all category counts
                    string chartQuery = "SELECT TM.category, COUNT(PM.hospitalNumber) AS CategoryCount " +
                                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                        "WHERE PM.currentWard = '" + safeWard + "' " +
                                        "GROUP BY TM.category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);

                                dbCounts[category] = count;
                            }
                        }
                    }

                    // *** FIX: Create a separate series for each category. This fixes the Legend and bar display. ***
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        // 1. Create a new series named after the category (This name appears in the legend!)
                        var categorySeries = chart1.Series.Add(categoryName);

                        // 2. Set chart properties for the new series
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor; // Set the color for the entire series

                        // 3. Get count from the database results, or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // 4. Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART FIX

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (NO CHANGE NEEDED HERE)
                    // ----------------------------------------------------

                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

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
                string com = "SELECT " +
                             "PM.hospitalNumber, PM.name, PM.surname, PM.gender, PM.currentWard, PM.isAdmitted, PM.admissionDate, PM.dischargeDate, " +
                             "TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                             "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                             "ORDER BY PM.hospitalNumber ASC";
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
                comboBox1.SelectedIndex = 0; 
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
    string query = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                   "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                   "WHERE ";
    switch (criteria)
    {
        case "Hospital Number":
            query += "PM.[hospitalNumber] = ?";
            break;
        case "Name":
            query += "PM.[name] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Surname":
            query += "PM.[surname] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Category":
            query += "TM.[category] LIKE ?";
            value = "%" + value + "%";
            break;
        case "Ward":
            query += "PM.[currentWard] LIKE ?";
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

        private void button5_Click(object sender, EventArgs e)
        {
            Reports obj = new Reports();
            this.Close();
            obj.Show();
        }

        private void overalStatisticsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ----------------------------------------------------
            // 0. CATEGORY CONSTANTS AND COLORS
            // ----------------------------------------------------
            const string CatAdmission = "Admission";
            const string CatTransferIn = "InterWardTransferIn";
            const string CatDischarge = "Discharge";
            const string CatTransferOut = "InterWardTransferOut";
            const string CatDeath = "Death";
            const string CatActivePatients = "CurrentInPatients"; // NEW CONSTANT FOR ACTIVE PATIENTS

            // Define a map for colors aligned with categories (for the chart)
            var categoryColors = new Dictionary<string, System.Drawing.Color>
{
    {CatActivePatients, System.Drawing.Color.DarkViolet}, // NEW COLOR MAPPING FOR CURRENT PATIENTS
    {CatAdmission, System.Drawing.Color.MediumSeaGreen},
    {CatTransferIn, System.Drawing.Color.SteelBlue},
    {CatDischarge, System.Drawing.Color.Firebrick},
    {CatTransferOut, System.Drawing.Color.Orange},
    {CatDeath, System.Drawing.Color.DarkRed}
};

            // --- Execution Start ---

            // Initializing visibility and clearing controls
            chart1.Visible = true;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView2.DataSource = null;

            // The previous code relating to ward selection (ToolStripMenuItem) has been removed.

            try
            {
                // Note: Assuming DatabaseHelper.ConnectionString is accessible and valid
                using (OleDbConnection con = new OleDbConnection(DatabaseHelper.ConnectionString))
                {
                    con.Open();

                    // ----------------------------------------------------
                    // 1. CHART SETUP AND UNIQUE PATIENT DATA LOADING (MOVEMENT CATEGORIES)
                    // ----------------------------------------------------

                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary (Unique Patients) for: ALL WARDS");
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count of Unique Patients";
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // *** OLEDB/ACCESS FIX: Use a Subquery to emulate COUNT(DISTINCT) ***
                    // Step 1: Subquery to select all DISTINCT patient/category pairs from the movement table.
                    string subQuery = "SELECT DISTINCT TM.category, TM.hospitalNumber " +
                                      "FROM tblPatientMovement AS TM";

                    // Step 2: Main query to count the results of the subquery, grouped by category.
                    string chartQuery = "SELECT Sub.category, COUNT(Sub.hospitalNumber) AS CategoryCount " +
                                        "FROM (" + subQuery + ") AS Sub " +
                                        "GROUP BY Sub.category";

                    // Step 3: Execute Query and store results in a map (dbCounts)
                    var dbCounts = new Dictionary<string, int>();

                    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                    {
                        using (OleDbDataReader reader = com.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int count = Convert.ToInt32(reader["CategoryCount"]);
                                dbCounts[category] = count;
                            }
                        }
                    }

                    // ----------------------------------------------------
                    // 1.5. CALCULATE TOTAL ACTIVE PATIENTS (NEW LOGIC)
                    // ----------------------------------------------------
                    // Query to find the count of unique patients whose latest movement is NOT Discharge or Death.
                    // This query finds the latest movement date for each patient and then checks the category of that latest movement.
                    string activePatientQuery =
                        "SELECT COUNT(T.hospitalNumber) AS ActiveCount " +
                        "FROM ( " +
                            "SELECT TM.hospitalNumber " +
                            "FROM tblPatientMovement AS TM " +
                            "INNER JOIN " +
                        // Sub-query to find the latest movement date for every patient
                                "(SELECT hospitalNumber, MAX(MovementDateTime) AS LatestDate " +
                                 "FROM tblPatientMovement " +
                                 "GROUP BY hospitalNumber) AS LastMove " +
                        // Join the movement table to the latest dates to get the category of the last movement
                            "ON TM.hospitalNumber = LastMove.hospitalNumber AND TM.MovementDateTime = LastMove.LatestDate " +
                        // Filter out Discharge and Death categories
                            "WHERE TM.category <> '" + CatDischarge + "' AND TM.category <> '" + CatDeath + "' " +
                            "GROUP BY TM.hospitalNumber" +
                        ") AS T";

                    using (OleDbCommand activeCmd = new OleDbCommand(activePatientQuery, con))
                    {
                        object result = activeCmd.ExecuteScalar();
                        // Store the current active count under the new category name
                        int activeCount = (result != DBNull.Value) ? Convert.ToInt32(result) : 0;
                        dbCounts[CatActivePatients] = activeCount;
                    }


                    // Step 4: Iterate through all predefined categories (including the new one) to plot the counts.
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;

                        // Get count from the database results (dbCounts), or default to 0
                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        // Add the single point.
                        categorySeries.Points.AddXY(categoryName, count);
                    }

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING (ALL RECORDS)
                    // ----------------------------------------------------

                    // DataGridView query now selects ALL records (no WHERE clause based on ward).
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);
                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;

                    con.Close();
                }
            }
            catch (Exception ex)
            {
                // Assuming MessageBox.Show is available in this context (e.g., Windows Forms)
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }
    }
}
