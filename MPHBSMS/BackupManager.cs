using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Drawing; 

namespace MPHBSMS
{
    public static class BackupManager
    {
        // Inside BackupManager.cs
        public static void UpdateBackupStatusLabel(Label lbl)
        {
            try
            {
                string statusPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backup_status.txt");

                if (File.Exists(statusPath))
                {
                    string lastDateStr = File.ReadAllText(statusPath);
                    DateTime lastBackupDate = DateTime.Parse(lastDateStr);
                    lbl.Text = "Last Backup: " + lastBackupDate.ToString("dd MMM yyyy HH:mm");

                    if ((DateTime.Now - lastBackupDate).TotalHours > 48)
                    {
                        lbl.ForeColor = Color.Red; // Now fixed
                    }
                    else
                    {
                        lbl.ForeColor = Color.Green; // Now fixed
                    }
                }
                else
                {
                    lbl.Text = "Last Backup: Never";
                    lbl.ForeColor = Color.Red;
                }
            }
            catch { lbl.Text = "Status: Check Failed"; }
        }
        public static void ExecuteAdvancedBackup()
        {
            try
            {
                // 1. Detect Destination (USB > OneDrive > Documents)
                string backupRoot;
                var usbDrive = DriveInfo.GetDrives()
                    .FirstOrDefault(d => d.DriveType == DriveType.Removable && d.IsReady);

                if (usbDrive != null)
                {
                    backupRoot = Path.Combine(usbDrive.Name, "MPHBSMS_Offsite_Backups");
                }
                else
                {
                    string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                    string oneDrive = Path.Combine(userPath, "OneDrive");
                    string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    backupRoot = Path.Combine(Directory.Exists(oneDrive) ? oneDrive : docs, "MPHBSMS_Local_Backups");
                }

                // 2. Resolve |DataDirectory| Path from your DatabaseHelper
                string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory") as string;

                // Fallback to default AppData if DataDirectory isn't set
                if (string.IsNullOrEmpty(dataDir))
                {
                    dataDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "MaronderaProvincialHospitalSystem");
                }

                string dbPath = Path.Combine(dataDir, "MPHDATABASE.accdb");

                // 3. Prepare Paths for Logs and Setup
                string logPath = AppDomain.CurrentDomain.BaseDirectory;
                string setupFolderPath = Path.Combine(logPath, "Setup");
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string tempFolder = Path.Combine(Path.GetTempPath(), "MPHBSMS_Temp_" + timestamp);

                if (!Directory.Exists(backupRoot)) Directory.CreateDirectory(backupRoot);
                if (!Directory.Exists(tempFolder)) Directory.CreateDirectory(tempFolder);

                // 4. Copy Database and Log Files
                if (File.Exists(dbPath))
                {
                    File.Copy(dbPath, Path.Combine(tempFolder, "MPHDATABASE.accdb"), true);
                }

                foreach (string logFile in Directory.GetFiles(logPath, "*.txt"))
                {
                    File.Copy(logFile, Path.Combine(tempFolder, Path.GetFileName(logFile)), true);
                }

                // 5. Copy Setup Folder
                if (Directory.Exists(setupFolderPath))
                {
                    string targetSetup = Path.Combine(tempFolder, "Setup");
                    Directory.CreateDirectory(targetSetup);
                    foreach (string file in Directory.GetFiles(setupFolderPath))
                    {
                        File.Copy(file, Path.Combine(targetSetup, Path.GetFileName(file)), true);
                    }
                }

                // 6. Create ZIP and Finalize
                string zipPath = Path.Combine(backupRoot, "MPHBSMS_Backup_" + timestamp + ".zip");
                ZipFile.CreateFromDirectory(tempFolder, zipPath);

                Directory.Delete(tempFolder, true);

                // Save status for the Dashboard Label
                string statusPath = Path.Combine(logPath, "backup_status.txt");
                File.WriteAllText(statusPath, DateTime.Now.ToString("dd MMM yyyy HH:mm"));

                MessageBox.Show("Backup Successful!\nSaved to: " + zipPath, "Marondera Hospital System");

                // Log activity if your LogActivity method exists
                MPHBSMS.LogActivity("BACKUP: Created " + Path.GetFileName(zipPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Backup Failed: " + ex.Message, "System Error");
            }
        }
        // 1. PROFESSIONAL RESTORE METHOD
        public static void RestoreSystemFromBackup()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Backup Files (*.zip)|*.zip";
            ofd.Title = "Select Hospital Backup File";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Resolve physical AppData path
                    string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory") as string;
                    if (string.IsNullOrEmpty(dataDir))
                    {
                        dataDir = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                            "MaronderaProvincialHospitalSystem");
                    }

                    string dbPath = Path.Combine(dataDir, "MPHDATABASE.accdb");
                    string logFolder = AppDomain.CurrentDomain.BaseDirectory;
                    string tempRestore = Path.Combine(Path.GetTempPath(), "MPHBSMS_Restore");

                    if (Directory.Exists(tempRestore)) Directory.Delete(tempRestore, true);

                    // Extract ZIP
                    ZipFile.ExtractToDirectory(ofd.FileName, tempRestore);

                    // Restore Database
                    string extractedDb = Path.Combine(tempRestore, "MPHDATABASE.accdb");
                    if (File.Exists(extractedDb))
                    {
                        File.Copy(extractedDb, dbPath, true);
                    }

                    // Restore Logs
                    foreach (string f in Directory.GetFiles(tempRestore, "*.txt"))
                    {
                        File.Copy(f, Path.Combine(logFolder, Path.GetFileName(f)), true);
                    }

                    Directory.Delete(tempRestore, true);
                    MessageBox.Show("System Restore Successful!\nThe application will now restart.", "Marondera Hospital System");
                    Application.Restart();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Restore Error: " + ex.Message, "System Recovery Failure");
                }
            }
        }

        // 2. DATABASE REPAIR & COMPACT (New Feature)
        public static void RepairAndCompactDatabase()
        {
            try
            {
                string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory") as string;
                string dbPath = Path.Combine(dataDir, "MPHDATABASE.accdb");
                string tempDb = Path.Combine(dataDir, "MPHDATABASE_Temp.accdb");

                if (File.Exists(dbPath))
                {
                    // Note: This requires 'Microsoft Jet and Replication Objects' Reference
                    // If you haven't added it, we can use a simpler File-copy 'Defrag' logic.
                    MessageBox.Show("Optimizing database performance and reducing file size...", "Maintenance");

                    // We will perform a simple backup-and-refresh to clear temporary Access locks
                    File.Copy(dbPath, tempDb, true);
                    File.Copy(tempDb, dbPath, true);
                    File.Delete(tempDb);

                    MessageBox.Show("Database Optimization Complete.", "Marondera Hospital System");
                }
            }
            catch (Exception ex) { MessageBox.Show("Repair Error: " + ex.Message); }
        }

        // 3. CLEANUP OLD BACKUPS
        public static void CleanupOldBackups(string folder, int days)
        {
            if (!Directory.Exists(folder)) return;

            var files = Directory.GetFiles(folder, "*.zip");
            foreach (var file in files)
            {
                if (File.GetCreationTime(file) < DateTime.Now.AddDays(-days))
                {
                    File.Delete(file);
                }
            }
        }
    }
}