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
                getColumnNameFromExcel(fileName);
                getRowsDataTypesFromExcelFile(fileName);
                GetDataFromExcelFile(fileName);
                Sql sql = new Sql(SheetName, ColumnNames, ColumnsDataTypes, dataTable);
                sql.ConvertExcelDataTypesToSql();
                sql.CreateTable();
                sql.CreateRows();
            }
        }

        public static void getColumnNameFromExcel(string fileName)
        {
            using (stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    {
                        reader.Read();
                        {
                            SheetName = reader.Name;
                            Console.WriteLine("Displaying Table Name:");
                            Console.WriteLine(SheetName);
                            var cols = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                            Console.Write("Displaying Column Names:");
                            foreach (var stuff in cols)
                            {
                                ColumnNames = new List<dynamic>();
                                ColumnNames = cols;
                                Console.WriteLine(stuff);
                                Console.ReadKey();
                            }
                        }
                    }
                }
            }
        }

        public static void getRowsDataTypesFromExcelFile(string fileName)
        {
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    reader.Read();
                    {
                        Console.WriteLine("Getting Columns Datatypes");
                        var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
                        ColumnsDataTypes = new List<dynamic>();
                        for (int i = 0; i < rows.Length; i++)
                        {
                            var COL = reader.GetFieldType(i);
                            ColumnsDataTypes.Add(COL);
                            Console.WriteLine(COL);
                        }
                        Console.WriteLine("Congrats!! Columns Datatypes displayed sucessfully");
                    }

                }
            }
        }
    }
}

