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
            // 1. Record the logout in the audit trail
            MPHBSMS.LogLogout();

            // 2. Clear the current user for security
            MPHBSMS.CurrentUser = "";
            
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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
// Assuming 'con' is an active OleDbConnection.

// *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
chart1.Titles.Clear();
chart1.Series.Clear();

chart1.Titles.Add("Ward Activity Summary for: " + ward);
chart1.ChartAreas[0].AxisX.Title = "Category";
chart1.ChartAreas[0].AxisY.Title = "Count";

chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

// FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                    "FROM ( " +
                        "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                        "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                        "WHERE PM.currentWard = '" + safeWard + "' " +
                    ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                    "GROUP BY Category";

// Step 1: Execute Query and store results in a map
var dbCounts = new Dictionary<string, int>();

try
{
    using (OleDbCommand com = new OleDbCommand(chartQuery, con))
    {
        using (OleDbDataReader reader = com.ExecuteReader())
        {
            while (reader.Read())
            {
                // Note: Reading the aliased column 'Category' instead of 'TM.category'
                string category = reader["Category"].ToString(); 
                int count = Convert.ToInt32(reader["CategoryCount"]);

                dbCounts[category] = count;
            }
        }
    }
}
catch (Exception ex)
{
    // Handle database errors (e.g., connection issues, query errors)
    MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    return; // Exit the method if data loading fails
}

// Step 2: Create a separate series for each category and populate the chart. (No change needed here)
foreach (var kvp in categoryColors)
{
    string categoryName = kvp.Key;
    System.Drawing.Color barColor = kvp.Value;

    var categorySeries = chart1.Series.Add(categoryName);
    categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
    categorySeries.IsValueShownAsLabel = true;
    categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
    categorySeries.Color = barColor;
    categorySeries.SetCustomProperty("PointWidth", "0.7");

    int count = 0;
    if (dbCounts.ContainsKey(categoryName))
    {
        count = dbCounts[categoryName];
    }

    categorySeries.Points.AddXY(categoryName, count);
}
// END OF CHART CODE

// DataGridView query with direct string injection
string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                  "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                  "WHERE PM.currentWard = '" + safeWard + "'";

OleDbCommand comm = new OleDbCommand(myquaery, con);

OleDbDataAdapter dm = new OleDbDataAdapter(comm);
DataTable dtt = new DataTable();
dm.Fill(dtt);
dataGridView2.DataSource = dtt;
                    

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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


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
                    // Assuming 'safeWard' is a checked and valid string variable holding the ward name.
                    // Assuming 'con' is an active OleDbConnection.

                    // *** 1. CHART INITIALIZATION AND CLEARING (Already correct) ***
                    chart1.Titles.Clear();
                    chart1.Series.Clear();

                    chart1.Titles.Add("Ward Activity Summary for: " + ward);
                    chart1.ChartAreas[0].AxisX.Title = "Category";
                    chart1.ChartAreas[0].AxisY.Title = "Count";

                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;

                    // FIX APPLIED TO CHART QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string chartQuery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                        "FROM ( " +
                                            "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                            "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                            "WHERE PM.currentWard = '" + safeWard + "' " +
                                        ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                        "GROUP BY Category";

                    // Step 1: Execute Query and store results in a map
                    var dbCounts = new Dictionary<string, int>();

                    try
                    {
                        using (OleDbCommand com = new OleDbCommand(chartQuery, con))
                        {
                            using (OleDbDataReader reader = com.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    // Note: Reading the aliased column 'Category' instead of 'TM.category'
                                    string category = reader["Category"].ToString();
                                    int count = Convert.ToInt32(reader["CategoryCount"]);

                                    dbCounts[category] = count;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors (e.g., connection issues, query errors)
                        MessageBox.Show("Error loading chart data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // Exit the method if data loading fails
                    }

                    // Step 2: Create a separate series for each category and populate the chart. (No change needed here)
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
                        categorySeries.Color = barColor;
                        categorySeries.SetCustomProperty("PointWidth", "0.7");

                        int count = 0;
                        if (dbCounts.ContainsKey(categoryName))
                        {
                            count = dbCounts[categoryName];
                        }

                        categorySeries.Points.AddXY(categoryName, count);
                    }
                    // END OF CHART CODE

                    // ----------------------------------------------------
                    // 2. DATAGRIDVIEW LOADING 
                    // ----------------------------------------------------

                    /*// FIX APPLIED TO DATAGRIDVIEW QUERY: Use Subquery to resolve COUNT(DISTINCT...) syntax error.
                    string myquaery = "SELECT Category, COUNT(HospitalID) AS CategoryCount " +
                                      "FROM ( " +
                                          "SELECT DISTINCT PM.hospitalNumber AS HospitalID, TM.category AS Category " +
                                          "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                          "WHERE PM.currentWard = '" + safeWard + "' " +
                                      ") AS FilteredMovements " + // CRITICAL: Subquery alias added
                                      "GROUP BY Category";

                    try
                    {
                        using (OleDbCommand comm = new OleDbCommand(myquaery, con))
                        {
                            OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                            DataTable dtt = new DataTable();
                            dm.Fill(dtt);
                            dataGridView2.DataSource = dtt;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle database errors for DataGridView loading
                        MessageBox.Show("Error loading DataGridView data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                    // DataGridView query with direct string injection
                    string myquaery = "SELECT PM.*, TM.MovementDateTime, TM.toWard, TM.fromWard, TM.category, TM.enteredBy " +
                                      "FROM tblPatientMaster AS PM INNER JOIN tblPatientMovement AS TM ON PM.hospitalNumber = TM.hospitalNumber " +
                                      "WHERE PM.currentWard = '" + safeWard + "'";

                    OleDbCommand comm = new OleDbCommand(myquaery, con);

                    OleDbDataAdapter dm = new OleDbDataAdapter(comm);
                    DataTable dtt = new DataTable();
                    dm.Fill(dtt);
                    dataGridView2.DataSource = dtt;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR!!\n" + ex.Message,
                    "Marondera Provincial Hospital",
                    MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void overallStatisticsToolStripMenuItem_Click(object sender, EventArgs e)
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
                    chart1.ChartAreas[0].AxisY.Title = "Unique Patients Counts";
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

                    // This query is inherently unreliable in OLEDB/Access due to the complex MAX/JOIN logic.
                    string activePatientQuery =
                        "SELECT COUNT(T.hospitalNumber) AS ActiveCount " +
                        "FROM ( " +
                            "SELECT TM.hospitalNumber " +
                            "FROM tblPatientMovement AS TM " +
                            "INNER JOIN " +
                                "(SELECT hospitalNumber, MAX(MovementDateTime) AS LatestDate " +
                                "FROM tblPatientMovement " +
                                "GROUP BY hospitalNumber) AS LastMove " +
                            "ON TM.hospitalNumber = LastMove.hospitalNumber AND TM.MovementDateTime = LastMove.LatestDate " +
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


                    // ----------------------------------------------------
                    // 1.6. CHART PLOTTING (FIXED TO RESOLVE COMPILATION ERRORS AND ADD SPACING)
                    // ----------------------------------------------------

                    // NEW FIX FOR SPACING: Modify Axis X settings to enable bar separation
                    chart1.ChartAreas[0].AxisX.IsLabelAutoFit = false;
                    chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -45; // Optional: Helps with long labels
                    chart1.ChartAreas[0].AxisX.IsMarginVisible = true;

                    // Step 4: Iterate through all predefined categories to plot the counts.
                    foreach (var kvp in categoryColors)
                    {
                        string categoryName = kvp.Key;
                        System.Drawing.Color barColor = kvp.Value;

                        var categorySeries = chart1.Series.Add(categoryName);
                        categorySeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                        categorySeries.IsValueShownAsLabel = true;
                        categorySeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;

                        // Set Color and SetCustomProperty on the Series object.
                        categorySeries.Color = barColor;

                        // ADJUSTED POINT WIDTH: Setting a smaller PointWidth value increases the space between bars.
                        categorySeries.SetCustomProperty("PointWidth", "0.5");

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
