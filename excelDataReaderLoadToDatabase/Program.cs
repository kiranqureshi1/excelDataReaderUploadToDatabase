using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelDataReaderConsoleApp
{
    class Program
    {
        private static List<string> listOfDatabases = new List<string>();
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public static void Main(string[] args)
        {
            string path = ConfigurationManager.AppSettings["excelSourcePath"].ToString();
            string search = ConfigurationManager.AppSettings["extension"].ToString();
            string database = ConfigurationManager.AppSettings["database"].ToString();
            Logger.Info("Starting validation");
            Logger.Info($"Looking for {search} files in {Regex.Unescape(path)}");

            ExcelDataReaderFile ExcelDataReaderFile = new ExcelDataReaderFile();
            string[] fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);

            Logger.Info($"Found {fileEntries.Length} file(s) to validate");

            dynamic counter = 0;
            BasicValidation BasicValidation = new BasicValidation();
            if (!BasicValidation.NoFilesInAFolder(fileEntries, path))
            {
                foreach (string file in fileEntries)
                {
                    BasicValidation basicValidation = new BasicValidation(file, 128, 32767, 128, 128);
                    counter += 1;

                    Logger.Info($"Validating {basicValidation.GetFileName(file)}");

                    if (basicValidation.InvalidColumnNames() || basicValidation.InvalidCellData())
                    {
                        //if the above statments are true then run those methods and log an error
                        Logger.Fatal($"{counter}) {basicValidation.GetFileName(file)} in {file} has failed validation. Check the error log for details.");
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                    }
                    else
                    {
                        ExcelDataReaderFile excelDataReaderFile = new ExcelDataReaderFile();
                        foreach (DataTable dataTable in excelDataReaderFile.GetDataTable(file))
                        {
                            Sql sql = new Sql(dataTable.TableName, dataTable, excelDataReaderFile.GetColumnNames(dataTable), excelDataReaderFile.RowsDataTypes(dataTable));
                            sql.UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(database);
                            Logger.Info($"{counter}) file number {counter} with file path {file} has been uploaded to the database");
                        }
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                    }
                }
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        public static void ChooseDatabase(Sql Sql)
        {
            ExcelDataReaderFile excelDataReaderFile = new ExcelDataReaderFile();
            if (DatabaseWithFolderNameExists(excelDataReaderFile.GetFolderName()))
            {
                listOfDatabases.Add(excelDataReaderFile.GetFolderName());
                Console.WriteLine($"would you like to use the existing database in the list below Yes/No");
                ExistingDatabases();
                var answer = Console.ReadLine();
                if (answer.ToLower() == "no")
                {
                    Console.WriteLine("What would you like to call your database?");
                    while (DatabaseWithFolderNameExists(Console.ReadLine()))
                    {
                        Console.WriteLine($"{Console.ReadLine()} already exists. The current list of databases are.");
                        ListOfDatabases();
                        Console.WriteLine("Please choose a different name");
                    }
                    if (Console.ReadLine() != null)
                    {
                        var newDb = Console.ReadLine();
                        Sql.CreateDbGet_Table_Columns_Rows_DataTypes(newDb);
                        listOfDatabases.Add(newDb);
                        Console.WriteLine("existing database");
                        ExistingDatabases();
                    }
                    else
                    {
                        Sql.UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(excelDataReaderFile.GetFolderName());
                    }
                }
                else
                {
                    Console.WriteLine("Choose database from the list below by typing its name.");
                    ExistingDatabases();
                    Sql.UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(Console.ReadLine());
                }
            }
            else
            {
                Sql.CreateDbGet_Table_Columns_Rows_DataTypes(excelDataReaderFile.GetFolderName());
            }
        }

        public static bool DatabaseWithFolderNameExists(string folderName)
        {
            string serverName = System.Configuration.ConfigurationManager.AppSettings["server"];
            var server = new Server(serverName);

            foreach (Database db in server.Databases)
            {
                if (db.Name.ToString() == folderName)
                {
                    return true;
                }
                else
                {
                }
            }
            return false;
        }

        public static void ListOfDatabases()
        {
            string serverName = System.Configuration.ConfigurationManager.AppSettings["server"];
            var server = new Server(serverName);
            foreach (Database db in server.Databases)
            {
                Console.WriteLine(db.Name);
            }
        }


        public static void ExistingDatabases()
        {
            foreach (string db in listOfDatabases)
            {
                Console.WriteLine(db);
            }
        }
    }
}
