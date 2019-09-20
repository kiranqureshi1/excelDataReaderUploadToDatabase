using ExcelDataReader;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReaderConsoleApp
{
    class program
    {
        public static string path = "C:\\Temp\\source";
        public static string search = "*.xlsx";
        public static String SheetName;
        public static List<dynamic> ColumnNames;
        public static List<dynamic> ColumnsDataTypes;
        public static DataTable dataTable;
        public static FileStream stream;
        public static string[] fileEntries;
        public static dynamic data;

        //public static void DataTypesFromExcel()
        public static void Main(string[] args)
        {
            fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);
            foreach (string fileName in fileEntries)
            {
                string[] fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);
                Sql sql = new Sql(SheetName, ColumnNames, ColumnsDataTypes, dataTable);
                sql.CreateTable();
            }
        }
    }
}

