using System;
using System.IO;
using System.Collections.Generic;

namespace MPHBSMS
{
    class MPHBSMS
    {
        public static string CurrentUser { get; set; }

        // Centralized path logic so it works on any computer
        private static string GetLogFolderPath()
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folder = Path.Combine(documentsPath, "Marondera Provincial Hospital Bed Statistics Reports", "SystemLogs");

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        public static void LogActivity(string activity)
        {
            try
            {
                string currentYear = DateTime.Now.Year.ToString();
                string fileName = "MPH_Log_" + currentYear + ".txt";

                // Saves to Documents\Marondera Provincial Hospital Bed Statistics Reports\SystemLogs
                string logPath = Path.Combine(GetLogFolderPath(), fileName);

                string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                string logEntry = "[" + timestamp + "] | User: " + (CurrentUser ?? "UNKNOWN").PadRight(12) + " | Action: " + activity + Environment.NewLine;

                File.AppendAllText(logPath, logEntry);
            }
            catch { }
        }

        // Standard method syntax for VS 2013 compatibility
        public static void LogLogin()
        {
            LogActivity("LOGIN: User successfully entered the system.");
        }

        public static void LogLogout()
        {
            LogActivity("LOGOUT: User session terminated/Application closed.");
        }

        public static void LogFailedLogin(string attemptedUsername)
        {
            LogActivity("SECURITY ALERT: Failed login attempt for username: " + attemptedUsername);
        }

        public static System.Collections.Generic.List<string> GetAvailableLogFiles()
        {
            System.Collections.Generic.List<string> myFiles = new System.Collections.Generic.List<string>();
            try
            {
                string path = GetLogFolderPath();
                string[] files = Directory.GetFiles(path, "MPH_Log_*.txt");

                foreach (string file in files)
                {
                    myFiles.Add(Path.GetFileName(file));
                }
            }
            catch { }
            return myFiles;
        }
    }
}