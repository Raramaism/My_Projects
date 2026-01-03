using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPHBSMS
{
    class MPHBSMS
    {
        public static string CurrentUser { get; set; }

        public static void LogActivity(string activity)
        {
            try
            {
                string currentYear = DateTime.Now.Year.ToString();
                string fileName = "MPH_Log_" + currentYear + ".txt";
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                string logEntry = "[" + timestamp + "] | User: " + (CurrentUser ?? "UNKNOWN").PadRight(12) + " | Action: " + activity + Environment.NewLine;
                File.AppendAllText(logPath, logEntry);
            }
            catch { }
        }

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
            string alert = "SECURITY ALERT: Failed login attempt for username: " + attemptedUsername;
            LogActivity(alert);
        }

        // Using the full name to avoid iTextSharp conflict
        public static System.Collections.Generic.List<string> GetAvailableLogFiles()
        {
            // Rename the variable 'logFiles' to 'myFiles' to fix Error 1
            System.Collections.Generic.List<string> myFiles = new System.Collections.Generic.List<string>();

            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory;

                // Find the files
                string[] files = Directory.GetFiles(path, "MPH_Log_*.txt");

                foreach (string file in files)
                {
                    // Add the filename to our renamed list
                    myFiles.Add(Path.GetFileName(file));
                }
            }
            catch
            {
                // Handle folder access errors if necessary
            }

            return myFiles; // Return the renamed list
           }
        }
    }

