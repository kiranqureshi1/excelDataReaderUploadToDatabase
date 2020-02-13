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
    public class ExcelDataReaderFile
    {
        public static string GetFolderName()
        {
            string path = ConfigurationManager.AppSettings["path"].ToString();
            // string lastFolderName = Path.GetFileName(Path.GetDirectoryName(path));
            DirectoryInfo dir_info = new DirectoryInfo(path);
            string directory = dir_info.Name;
            return directory;
        }

        public static dynamic getSheetNameFromExcel(string fileName)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    {
                        reader.Read();
                        {
                            string SheetName = reader.Name;
                            return SheetName;
                        }
                    }
                }
            }
        }


        public static dynamic getColumnNameFromExcel(string fileName)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    {
                        reader.Read();
                        {
                            var SheetName = reader.Name;
                            var columnNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                            return columnNames;
                        }
                    }
                }
            }
        }

        public static dynamic getRowsDataTypesFromExcelFile(string fileName)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    reader.Read();
                    {
                        var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
                        List<dynamic> ColumnsDataTypes = new List<dynamic>();
                        for (int i = 0; i < rows.Length; i++)
                        {
                            var COL = reader.GetFieldType(i);
                            ColumnsDataTypes.Add(COL);
                        }
                        return ColumnsDataTypes;
                    }

                }
            }
        }

        public static dynamic GetDataTableExcelFile(string fileName)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                var dataSet = reader.AsDataSet(conf);
                DataTable dataTable = dataSet.Tables[0];
                for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (var j = 0; j < dataTable.Columns.Count; j++)
                    {
                        dynamic data = dataTable.Rows[i][j];
                    }
                }
                return dataTable;
            }
        }
    }
}
