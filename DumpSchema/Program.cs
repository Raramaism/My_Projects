using System;
using System.IO;
using System.Text;
using MPHBSMS;

namespace DumpSchema
{
    class Program
    {
        static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;

                // If your DatabaseHelper requires initialization, you can do it here:
                try
                {
                    // DatabaseHelper.InitializeDatabase(); // uncomment if needed
                    // Or set the connection string explicitly:
                    // DatabaseHelper.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\to\your\database.accdb;";
                }
                catch (Exception exInit)
                {
                    Console.WriteLine("Warning: DatabaseHelper initialization threw: " + exInit.Message);
                }

                var rm = new ReportManager();

                Console.WriteLine("Dumping schema for tblPatientMovement...");
                string movementSchema = rm.DumpTableSchema("tblPatientMovement");
                Console.WriteLine(movementSchema);

                Console.WriteLine();
                Console.WriteLine("Dumping schema for tblPatientMaster...");
                string masterSchema = rm.DumpTableSchema("tblPatientMaster");
                Console.WriteLine(masterSchema);

                Console.WriteLine();
                Console.WriteLine("Dumping schema for userAccounts...");
                string usersSchema = rm.DumpTableSchema("userAccounts");
                Console.WriteLine(usersSchema);

                string output = movementSchema + Environment.NewLine + masterSchema + Environment.NewLine + usersSchema;
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string outPath = Path.Combine(docs, "dump_schema.txt");
                File.WriteAllText(outPath, output, Encoding.UTF8);

                Console.WriteLine();
                Console.WriteLine("Schema written to: " + outPath);
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error while dumping schema: " + ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return 1;
            }
        }
    }
}