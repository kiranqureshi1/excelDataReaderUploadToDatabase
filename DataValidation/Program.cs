using ExcelDataReader;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReaderConsoleApp
{
    class program
    {

        public static void Main(string[] args)
        {
            string path = ConfigurationManager.AppSettings["path"].ToString();
            string search = ConfigurationManager.AppSettings["search"].ToString();
            Console.WriteLine("starting the app...");
            Console.WriteLine($"Path: {path}");
            Console.WriteLine($"file extension: {search}");
            // Console.WriteLine(ExcelDataReaderFile.GetFolderName());
            BasicValidation basicValidation = new BasicValidation();
            string[] fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);
            Console.WriteLine($"There are {fileEntries.Length} in {path}");
            Console.ReadKey();
            dynamic counter = 0;
            if (!basicValidation.NoFilesInAFolder(fileEntries, path))
            //{
            //}
            //else
            {
                foreach (string fileName in fileEntries)
                {
                    BasicValidation BasicValidation = new BasicValidation(fileName);
                    counter += 1;
                    Console.ForegroundColor = ConsoleColor.Red;
                    BasicValidation.SheetNameValidation();
                    BasicValidation.FileNameValidation();
                    BasicValidation.InvalidColumnNames();
                    BasicValidation.InvalidCellData();
                    LogWriter log = new LogWriter($"{counter}) so file number {counter} with file path {fileName}, cant be uploaded to the database", "Review the file for the above issues");
                    log.LogWrite();
                    Console.WriteLine($"{counter}) There was a problem with file number {counter}, so file with file path {fileName}, cant be uploaded to the database");
                    Console.WriteLine("--------------------------");
                    Console.ReadKey();
                }
            }
        }
    }
}


