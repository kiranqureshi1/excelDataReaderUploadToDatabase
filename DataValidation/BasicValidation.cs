using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReaderConsoleApp
{
    public class BasicValidation
    {
        private string fileName;

        public BasicValidation()
        {
        }

        public BasicValidation(string FileName)
        {
            fileName = FileName;
        }
        public Boolean NoFilesInAFolder(string[] fileEntries, string path)
        {
            if (fileEntries.Length == 0)
            {
                LogWriter log = new LogWriter($"There are no files in the folder {path}", "Review the folder");
                log.LogWrite();
                Console.WriteLine($"There are no files in the folder {path}");
                Console.ReadKey();
                return true;
            }
            else
            {
                return false;
            }
        }


        public bool SheetNameValidation()
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    {
                        reader.Read();
                        {
                            string SheetName = reader.Name;
                            if (SheetName.Length > 128)
                            {
                                LogWriter log = new LogWriter($" FileName: {SheetName} exceeds 128 limit", "if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
                                log.LogWrite();
                                Console.WriteLine($" FileName: {SheetName} exceeds 128 limit, if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
            }
        }


        public bool FileNameValidation()
        {
            if (fileName.Length > 128)
            {
                LogWriter log = new LogWriter($"{fileName} exceeds 128 characters", "Sql doesnt take files with file name size longer than 128 characters");
                log.LogWrite();
                Console.WriteLine($"{fileName} exceeds 128 characters, Sql doesnt take files with file name size longer than 128 characters");
                return true;
            }
            else
            {
                return false;
            }
        }


        public Boolean InvalidColumnNames()
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    reader.Read();
                    {
                        int counter = 0;
                        var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                        if (ColumnsNames.Count != 0 && reader.Read() == true)
                        {
                            foreach (var columnName in ColumnsNames)
                            {
                                counter += 1;
                                if (columnName != null && columnName.ToString().Length > 128)
                                {
                                    LogWriter log = new LogWriter($"Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name}. ColumnName: {columnName.ToString()}", "if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction");
                                    log.LogWrite();
                                    //Console.WriteLine(i.ToString().Length);
                                    Console.WriteLine($" Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name } if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction.");
                                    Console.WriteLine($"Column Name: {columnName.ToString()}");
                                    //Console.ReadKey();
                                    return true;
                                }
                                else if (columnName == null)
                                {
                                    LogWriter log = new LogWriter($"Column number {counter} is missing in fileName: {reader.Name}", "if the file doesnt have column names, Sql doesnt upload such files.");
                                    log.LogWrite();
                                    Console.WriteLine($" Column number {counter} is missing in fileName: {reader.Name}");
                                    //Console.ReadKey();
                                    return true;
                                }
                                else
                                {
                                    //Console.WriteLine(reader.Name);
                                    //Console.WriteLine($" Column number {counter} has acceptable size in fileName: {reader.Name}, ColumnName: {columnName.ToString()}");
                                    //Console.ReadKey();
                                    //return false;
                                }
                            }
                        }
                        else
                        {
                            LogWriter log = new LogWriter($" FileName: {reader.Name} is empty", "if the file is empty, Sql doesnt upload such files.");
                            log.LogWrite();
                            Console.WriteLine($" FileName: {reader.Name} is empty");
                            //Console.ReadKey();
                            return true;
                        };
                    }
                    return false;
                }
            }
        }

        public bool InvalidCellData()
        {
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
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
                var dataTable = dataSet.Tables[0];
                var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
                for (var ColumnIndexNumber = 0; ColumnIndexNumber < dataTable.Columns.Count; ColumnIndexNumber++)
                {
                    for (var RowIndexNumber = 0; RowIndexNumber < dataTable.Rows.Count; RowIndexNumber++)
                    {
                        //RowIndexNumber is row number starts from index number 0
                        //ColumnIndexNumber is column number starts from index number 0
                        var data = dataTable.Rows[RowIndexNumber][ColumnIndexNumber];
                        if (data.GetType() == reader.GetFieldType(ColumnIndexNumber) && data.ToString().Length > 128)
                        {
                            LogWriter log = new LogWriter($" FileName: {reader.Name} has data in a row {RowIndexNumber} column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data } ", "if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
                            log.LogWrite();
                            Console.WriteLine($" FileName: {reader.Name} has data in a row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data }. if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
                            //Console.ReadKey();
                            return true;
                        }
                        else if (data.ToString().Length != 0 & data.ToString().Length <= 128 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                        {
                            LogWriter log = new LogWriter($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }", "if data in a cell has wrong dataType, Sql doesnt upload such files.");
                            log.LogWrite();
                            Console.WriteLine($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }. If data in a cell has wrong dataType, Sql doesnt upload such files.");
                            //Console.ReadKey();
                            return true;
                        }
                        else if (data.ToString().Length == 0 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                        {
                            LogWriter log = new LogWriter($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}.", "if a file has an empty cell, Sql doesnt upload such files.");
                            log.LogWrite();
                            Console.WriteLine($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}. If a file has an empty cell, Sql doesnt upload such files.");
                            //Console.ReadKey();
                            return true;
                        }
                        else if (data.ToString().Length > 128 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                        {
                            LogWriter log = new LogWriter($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}", "if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
                            log.LogWrite();
                            Console.WriteLine($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}", "if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
                            //Console.ReadKey();
                            return true;
                        }
                        else
                        {
                            //Console.WriteLine($"{data} good size and  matches {data.GetType()}");
                            //Console.ReadKey();
                        }
                    }
                }
                return false;
            }
        }
    }
}
