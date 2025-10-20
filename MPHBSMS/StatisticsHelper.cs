using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms; // Required to pass controls for UI update

public static class StatisticsHelper
{
    /// <summary>
    /// Calculates all 7 ward statistics for a selected ward and updates the UI labels.
    /// </summary>
    public static void UpdateWardStatistics(string selectedWard, OleDbConnection con, Form callingForm)
    {
        // 1. INPUT VALIDATION
        if (string.IsNullOrWhiteSpace(selectedWard))
        {
            MessageBox.Show("Ward name cannot be empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            // Ensure connection is open (Caller should manage connection, but we check here for safety)
            if (con.State != System.Data.ConnectionState.Open) con.Open();

            // --- QUERY 1: CURRENT TOTAL BEDS OCCUPIED (Stat 7) ---
            string bedsOccupiedQuery = @"
                SELECT COUNT([hospitalNumber]) 
                FROM tblPatientMaster 
                WHERE IsAdmitted = 'Yes' AND CurrentWard = ?";

            int totalBedsOccupied;
            using (OleDbCommand bedsCmd = new OleDbCommand(bedsOccupiedQuery, con))
            {
                bedsCmd.Parameters.Add("?", OleDbType.VarChar).Value = selectedWard;
                totalBedsOccupied = Convert.ToInt32(bedsCmd.ExecuteScalar());
            }

            // --- QUERY 2: TODAY'S MOVEMENT COUNTS (Stats 2, 3, 4, 5, 6) ---
            string today = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");

            string movementQuery = @"
                SELECT category, COUNT(movementID) AS Total
                FROM tblPatientMovement 
                WHERE 
                    (ToWard = ? OR FromWard = ?) 
                    AND MovementDateTime >= ?
                GROUP BY category";

            Dictionary<string, int> movementCounts = new Dictionary<string, int>
            {
                {"Admission", 0}, {"Transfer In", 0}, {"Transfer Out", 0}, 
                {"Discharge", 0}, {"Death", 0}
            };

            using (OleDbCommand moveCmd = new OleDbCommand(movementQuery, con))
            {
                moveCmd.Parameters.Add("?", OleDbType.VarChar).Value = selectedWard;
                moveCmd.Parameters.Add("?", OleDbType.VarChar).Value = selectedWard;
                moveCmd.Parameters.Add("?", OleDbType.VarChar).Value = today;

                using (OleDbDataReader reader = moveCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string category = reader["category"].ToString();
                        int total = Convert.ToInt32(reader["Total"]);
                        if (movementCounts.ContainsKey(category))
                        {
                            movementCounts[category] = total;
                        }
                    }
                }
            }

            // 3. CALCULATE AND UPDATE UI (All 7 Labels)

            int totalAdmissions = movementCounts["Admission"];
            int totalTransferIn = movementCounts["Transfer In"];
            int totalDischarges = movementCounts["Discharge"];
            int totalDeaths = movementCounts["Death"];

            int broughtForward = totalBedsOccupied - (totalAdmissions + totalTransferIn);

            // --- UPDATE LABELS ---
            // This requires casting the generic Form object back to your specific Menu form type
            // (assuming your form is named 'Menu' or similar, you may need to adjust this cast)
            var menuForm = callingForm as Menu;

            if (menuForm != null)
            {
                // 1. Brought forword
                menuForm.lblBroughtForward.Text = broughtForward.ToString("D2");

                // 2. Total Admissions
                menuForm.lblTotalAdmissions.Text = totalAdmissions.ToString("D2");

                // 3. Total Transfer in (First listing)
                menuForm.lblTotalTransferIn1.Text = totalTransferIn.ToString("D2");

                // 4. Total discharges
                menuForm.lblTotalDischarges.Text = totalDischarges.ToString("D2");

                // 5. Total Transfer in (Second listing)
                menuForm.lblTotalTransferIn2.Text = totalTransferIn.ToString("D2");

                // 6. Total deaths
                menuForm.lblTotalDeaths.Text = totalDeaths.ToString("D2");

                // 7. Total beds occupied
                menuForm.lblTotalBedsOccupied.Text = totalBedsOccupied.ToString("D2");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error loading ward statistics: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}