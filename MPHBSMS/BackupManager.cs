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
        // ... (UpdateBackupStatusLabel and ExecuteAdvancedBackup remain as we fixed them)

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
                    lbl.ForeColor = (DateTime.Now - lastBackupDate).TotalHours > 48 ? Color.Red : Color.Green;
                }
                else
                {
                    lbl.Text = "Last Backup: Never";
                    lbl.ForeColor = Color.Red;
                }
            }
            catch { lbl.Text = "Status: Check Failed"; }
        }

        public static void ExecuteAdvancedBackup(bool isClosing = false)
        {
            try
            {
                string backupRoot;
                var usbDrive = DriveInfo.GetDrives().FirstOrDefault(d => d.DriveType == DriveType.Removable && d.IsReady);

                if (usbDrive != null)
                {
                    DialogResult result = MessageBox.Show("USB Drive (" + usbDrive.Name + ") detected. Create backup now?", "External Backup", MessageBoxButtons.YesNo);
                    if (result == DialogResult.No) return;
                    backupRoot = Path.Combine(usbDrive.Name, "MPHBSMS_Offsite_Backups");
                }
                else
                {
                    string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    backupRoot = Path.Combine(docs, "MPHBSMS_Local_Backups");
                }

                string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory") as string;
                if (string.IsNullOrEmpty(dataDir)) dataDir = AppDomain.CurrentDomain.BaseDirectory;

                string dbPath = Path.Combine(dataDir, "MPHDATABASE.accdb");
                string logPath = AppDomain.CurrentDomain.BaseDirectory;
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string tempFolder = Path.Combine(Path.GetTempPath(), "MPH_Temp_" + timestamp);

                if (!Directory.Exists(backupRoot)) Directory.CreateDirectory(backupRoot);
                if (!Directory.Exists(tempFolder)) Directory.CreateDirectory(tempFolder);

                if (File.Exists(dbPath)) File.Copy(dbPath, Path.Combine(tempFolder, "MPHDATABASE.accdb"), true);
                foreach (string f in Directory.GetFiles(logPath, "*.txt")) File.Copy(f, Path.Combine(tempFolder, Path.GetFileName(f)), true);

                string zipPath = Path.Combine(backupRoot, "MPH_Backup_" + timestamp + ".zip");
                ZipFile.CreateFromDirectory(tempFolder, zipPath);
                Directory.Delete(tempFolder, true);
                File.WriteAllText(Path.Combine(logPath, "backup_status.txt"), DateTime.Now.ToString("dd MMM yyyy HH:mm"));

                
            }
            catch (Exception ex) { if (!isClosing) MessageBox.Show("Backup Error: " + ex.Message); }
        }

        // --- RESTORED METHOD START ---
        public static void RestoreSystemFromBackup()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Backup Files (*.zip)|*.zip";
            ofd.Title = "Select Hospital Backup File";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory") as string;
                    if (string.IsNullOrEmpty(dataDir)) dataDir = AppDomain.CurrentDomain.BaseDirectory;

                    string dbPath = Path.Combine(dataDir, "MPHDATABASE.accdb");
                    string logFolder = AppDomain.CurrentDomain.BaseDirectory;
                    string tempRestore = Path.Combine(Path.GetTempPath(), "MPH_Restore");

                    if (Directory.Exists(tempRestore)) Directory.Delete(tempRestore, true);
                    ZipFile.ExtractToDirectory(ofd.FileName, tempRestore);

                    // Restore Database
                    string extractedDb = Path.Combine(tempRestore, "MPHDATABASE.accdb");
                    if (File.Exists(extractedDb)) File.Copy(extractedDb, dbPath, true);

                    // Restore Logs
                    foreach (string f in Directory.GetFiles(tempRestore, "*.txt"))
                    {
                        File.Copy(f, Path.Combine(logFolder, Path.GetFileName(f)), true);
                    }

                    Directory.Delete(tempRestore, true);
                    MessageBox.Show("Restore Successful! Restarting system...");
                    Application.Restart();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Restore Error: " + ex.Message);
                }
            }
        }
        // --- RESTORED METHOD END ---
    }
}