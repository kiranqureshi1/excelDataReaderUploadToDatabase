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
        // private string fileName;
        public IExcelDataReader reader;

        public string GetFolderName()
        {
            string path = ConfigurationManager.AppSettings["path"].ToString();
            // string lastFolderName = Path.GetFileName(Path.GetDirectoryName(path));
            DirectoryInfo dir_info = new DirectoryInfo(path);
            string directory = dir_info.Name;
            return directory;
        }

        public dynamic GetDataTable(string fileName)
        {
            using (var reader = ExcelReaderFactory.CreateReader(File.Open(fileName, FileMode.Open, FileAccess.Read)))
            {
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                var dataSet = reader.AsDataSet(conf);
                var table = new DataTable();
                //Console.WriteLine(dataSet.DataSetName);
                return dataSet.Tables;
            }
        }



        public dynamic GetColumnNames(DataTable dataTable)
        {
            List<dynamic> columns = new List<dynamic>();
            //foreach (dynamic column in dataTable.Columns)
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var COL = dataTable.Columns[i];
                //Console.WriteLine(COL.ColumnName);
                columns.Add(COL.ColumnName);
            }
            return columns;
        }

        public dynamic RowsDataTypes(DataTable dataTable)
        {
            //GetDataTableExcelFile().Columns.GetType();
            List<dynamic> ColumnsDataTypes = new List<dynamic>();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var COL = dataTable.Columns[i];
                ColumnsDataTypes.Add(COL.DataType);
            }
            foreach (dynamic dataType in ColumnsDataTypes)
            {
                //Console.WriteLine(dataType);
            }
            return ColumnsDataTypes;
        }

        public string GetAddress(int column, int row)
        {
            column++;
            StringBuilder stringbuilder = new StringBuilder();
            do
            {
                column--;
                stringbuilder.Insert(0, (char)('A' + (column % 26)));
                column /= 26;

            } while (column > 0);
            stringbuilder.Append(row + 1);
            return stringbuilder.ToString();
        }


        //public ExcelDataReaderFile()
        //{
        //}
        //public ExcelDataReaderFile(string FileName)
        //{
        //    fileName = FileName;
        //    // reader = ExcelReaderFactory.CreateReader(File.Open(fileName, FileMode.Open, FileAccess.Read));
        //    // reader = Reader;
        //}

        //public IExcelDataReader StreamAndReadExcelFile()
        //{
        //    reader = ExcelReaderFactory.CreateReader(File.Open(fileName, FileMode.Open, FileAccess.Read));
        //    return reader;
        //}

        //public dynamic getSheetNameFromExcel()
        //{
        //    using (StreamAndReadExcelFile())
        //    // using(reader)
        //    {
        //        {
        //            reader.Read();
        //            {
        //                string SheetName = reader.Name;
        //                return SheetName;
        //            }
        //        }
        //    }
        //}


        //public dynamic getColumnNameFromExcel()
        //{
        //    //using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
        //    //{
        //    //    using (var reader = ExcelReaderFactory.CreateReader(stream))
        //    using (StreamAndReadExcelFile())
        //    //  using (reader)
        //    {
        //        reader.Read();
        //        {
        //            var SheetName = reader.Name;
        //            var columnNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
        //            return columnNames;
        //        }
        //    }
        //    //}
        //}

        //public dynamic getRowsDataTypesFromExcelFile()
        //{
        //    using (StreamAndReadExcelFile())
        //    //using(reader)
        //    {
        //        reader.Read();
        //        {
        //            var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
        //            List<dynamic> ColumnsDataTypes = new List<dynamic>();
        //            for (int i = 0; i < rows.Length; i++)
        //            {
        //                var COL = reader.GetFieldType(i);
        //                ColumnsDataTypes.Add(COL);
        //            }
        //            return ColumnsDataTypes;
        //        }

        //    }
        //    // }
        //}

        //public DataTable GetDataTableExcelFile()
        //{
        //    StreamAndReadExcelFile();
        //    // using (reader)
        //    // { 
        //    var conf = new ExcelDataSetConfiguration
        //    {
        //        ConfigureDataTable = _ => new ExcelDataTableConfiguration
        //        {
        //            UseHeaderRow = true
        //        }
        //    };
        //    var dataSet = reader.AsDataSet(conf);
        //    var table = new DataTable();
        //    foreach(DataTable dataTable in dataSet.Tables)
        //    {
        //       // DataTable dataTable = dataSet.Tables[0];
        //        for (var i = 0; i < dataTable.Rows.Count; i++)
        //        {
        //            for (var j = 0; j < dataTable.Columns.Count; j++)
        //            {
        //                dynamic data = dataTable.Rows[i][j];
        //            }
        //        }
        //        Console.WriteLine(dataTable);
        //        table = dataTable;
        //    }
        //    Console.WriteLine(table);
        //    return table;
        //}
    }
}
