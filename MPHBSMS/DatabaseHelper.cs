using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MPHBSMS
{
   
    class DatabaseHelper
    {

        public static string ConnectionString
        {
            get
            {
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MPHDATABASE.accdb;Persist Security Info=False;";
            }
        }

        public static void InitializeDatabase()
        {
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MaronderaProvincialHospitalSystem");

            if (!Directory.Exists(appDataPath))
                Directory.CreateDirectory(appDataPath);

            string sourceDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MPHDATABASE.accdb");
            string targetDb = Path.Combine(appDataPath, "MPHDATABASE.accdb");

            if (!File.Exists(targetDb))
            {
                File.Copy(sourceDb, targetDb);
            }

            AppDomain.CurrentDomain.SetData("DataDirectory", appDataPath);

        }
    }
}
