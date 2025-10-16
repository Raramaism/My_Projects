using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MPHBSMS
{
    /*

    class DatabaseHelper
    {
        public static string DatabasePath { get; private set; }

        public static string ConnectionString
        {
            get
            {
                //return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DatabasePath;
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DatabasePath + ";Persist Security Info=False;";

            }
        }

       
        public static void InitializeDatabase()
        {
            string appDataPath = Path.Combine(
                 Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MaronderaProvincialHospitalSystem");

            if (!Directory.Exists(appDataPath))
                Directory.CreateDirectory(appDataPath);

            string sourceDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MPHBSMS.accdb");
            string targetDb = Path.Combine(appDataPath, "MPHBSMS.accdb");

            if (!File.Exists(targetDb))
            {
                File.Copy(sourceDb, targetDb);
            }

            DatabasePath = targetDb;
        }
    }*/

    class DatabaseHelper
    {
        // 1. Remove DatabasePath field, as it's not needed with |DataDirectory|

        public static string ConnectionString
        {
            get
            {
                // 2. Use the new connection string format
                // |DataDirectory| will be substituted by the framework 
                // with the path set in AppDomain.CurrentDomain.SetData
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MPHBSMS.accdb;Persist Security Info=False;";
            }
        }

        public static void InitializeDatabase()
        {
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MaronderaProvincialHospitalSystem");

            if (!Directory.Exists(appDataPath))
                Directory.CreateDirectory(appDataPath);

            string sourceDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MPHBSMS.accdb");
            string targetDb = Path.Combine(appDataPath, "MPHBSMS.accdb");

            if (!File.Exists(targetDb))
            {
                File.Copy(sourceDb, targetDb);
            }

            // 3. CRITICAL: Set the DataDirectory path for the |DataDirectory| substitution string
            AppDomain.CurrentDomain.SetData("DataDirectory", appDataPath);

            // 4. Remove DatabasePath = targetDb; (no longer needed)
        }
    }
}
