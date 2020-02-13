using ExcelDataReader;
using ExcelDataReaderConsoleApp;
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
        private readonly string file;
        private IExcelDataReader reader;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private readonly int columnNameSizeLimit;
        private readonly int rowDataSizeLimit;
        private readonly int sheetNameSizeLimit;
        private readonly int fileNameSizeLimit;
        private bool errorDetected;
        private dynamic ColumnAlreadyMatchedA;
        private dynamic ColumnAlreadyMatchedB;

        public BasicValidation()
        { }
        public BasicValidation(string FileName, int ColumnNamseSizeLimit, int RowDataSizeLimit, int SheetNameSizeLimit, int FileNameSizeLimit)
        {
            file = FileName;
            columnNameSizeLimit = ColumnNamseSizeLimit;
            rowDataSizeLimit = RowDataSizeLimit;
            sheetNameSizeLimit = SheetNameSizeLimit;
            fileNameSizeLimit = FileNameSizeLimit;
        }

        public string GetFileName(string file)
        {
            string fileName = Path.GetFileName(Path.GetFileName(file));
            return fileName;
        }

        public IExcelDataReader ReadAndStreamFile()
        {
            reader = ExcelReaderFactory.CreateReader(File.Open(file, FileMode.Open, FileAccess.Read));
            return reader;
        }

        public Boolean NoFilesInAFolder(string[] fileEntries, string path)
        {
            if (fileEntries.Length == 0)
            {
                Logger.Error($"There are no files in the folder {path} to validate. Review the folder.");
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool SheetNameValidation()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    string SheetName = reader.Name;
                    if (SheetName.Length > sheetNameSizeLimit)
                    {
                        Logger.Error($"[{GetFileName(file)}]{SheetName} exceeds {sheetNameSizeLimit} character sheet name limit. Supply a valid sheet name.");
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }


        public bool FileNameValidation()
        {
            if (GetFileName(file).Length > fileNameSizeLimit)
            {
                Logger.Error($"{GetFileName(file)} exceeds {fileNameSizeLimit} character file name limit. Supply a valid file name.");
                return true;
            }
            else
            {
                return false;
            }
        }

        public Boolean DuplicateColumnNames()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                    var excel = new ExcelDataReaderFile();
                    errorDetected = false;
                    //looping through column names
                    for (int columnIndexNumber = 0; columnIndexNumber < ColumnsNames.Count; columnIndexNumber++)
                    {
                        //looping through rows
                        for (int columnIndexNum = 0; columnIndexNum < ColumnsNames.Count; columnIndexNum++)
                        {
                            var cellAddressA = excel.GetAddress(columnIndexNumber, 0);
                            var cellAddressB = excel.GetAddress(columnIndexNum, 0);
                            //so lets say we are talking about column A and Column C, if ColumnA and Column C are not null 
                            if (ColumnsNames[columnIndexNumber] != null && ColumnsNames[columnIndexNum] != null)
                            {
                                //so if column A were never put against column C to check if they match 
                                if (ColumnsNames[columnIndexNumber] != ColumnAlreadyMatchedA && ColumnsNames[columnIndexNum] != ColumnAlreadyMatchedB)
                                {
                                    //check every column against eachother apart from checking it with itself ( say if columnIndexNumber and columnIndexNum are the same then it emans we are trying to match it with itself)
                                    // any column say ColumnsNames[columnIndexNumber] is column A already matched with B, C and D so going backwards doesnt check column B,C,D against column A
                                    if (ColumnsNames[columnIndexNumber] == ColumnsNames[columnIndexNum] && columnIndexNumber != columnIndexNum)
                                    {
                                        // and so Column a is given  "ColumnAlreadyMatchedA"name
                                        ColumnAlreadyMatchedA = ColumnsNames[columnIndexNumber];
                                        ColumnAlreadyMatchedB = ColumnsNames[columnIndexNum];
                                        Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddressA} with column name {ColumnsNames[columnIndexNumber]} matches {cellAddressB} with column name {ColumnsNames[columnIndexNum]}");
                                        errorDetected = true;
                                    }
                                    else
                                    {
                                        // errorDetected = false;
                                    }
                                }
                            }
                        }
                    }
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }
        }


        public Boolean InvalidColumnNames()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    int counter = 0;
                    var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                    if (ColumnsNames.Count != 0 && reader.Read() == true)
                    {
                        errorDetected = false;
                        for (int columnNumber = 0; columnNumber < ColumnsNames.Count; columnNumber++)
                        {
                            var excel = new ExcelDataReaderFile();
                            var cellAddress = excel.GetAddress(counter, 0);
                            counter += 1;
                            if (ColumnsNames[columnNumber] != null && ColumnsNames[columnNumber].ToString().Length > columnNameSizeLimit)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {columnNumber.ToString().Length} characters long and exceeds {columnNameSizeLimit} character column name limit. Supply a valid column name.");
                                errorDetected = true;
                            }
                            else if (ColumnsNames[columnNumber] == null)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is empty. Supply a valid column name.");
                                errorDetected = true;
                            }
                            else
                            {
                            }
                            continue;
                        }
                    }
                    else
                    {
                        Logger.Error($"[{GetFileName(file)}]{reader.Name} is empty and cannot be validated. Supply a non-empty file.");
                        errorDetected = true;
                    };
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }
        }


        public bool InvalidCellData()
        {
            using (ReadAndStreamFile())
            {
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
                errorDetected = false;
                for (var ColumnIndexNumber = 0; ColumnIndexNumber < dataTable.Columns.Count; ColumnIndexNumber++)
                {
                    for (var RowIndexNumber = 0; RowIndexNumber < dataTable.Rows.Count; RowIndexNumber++)
                    {
                        //RowIndexNumber is row number starts from index number 0
                        //ColumnIndexNumber is column number starts from index number 0
                        var data = dataTable.Rows[RowIndexNumber][ColumnIndexNumber];
                        var excel = new ExcelDataReaderFile();
                        var cellAddress = excel.GetAddress(ColumnIndexNumber, RowIndexNumber + 1);
                        if (data.ToString().Length != 0)
                        {
                            if (data.GetType() == reader.GetFieldType(ColumnIndexNumber) && data.ToString().Length > rowDataSizeLimit)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {data.ToString().Length} characters long and exceeds {rowDataSizeLimit} character cell contents limit. Supply valid cell contents.");
                                errorDetected = true;
                            }
                            else if (data.ToString().Length <= rowDataSizeLimit & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} data {data} data type {data.GetType()} does not match data type of column data {reader.GetFieldType(ColumnIndexNumber)}. Supply data with a consistent data type.");
                                errorDetected = true;
                            }
                            else if (data.ToString().Length > rowDataSizeLimit & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {data.ToString().Length} characters long and exceeds {rowDataSizeLimit} character cell contents limit. Supply valid cell contents. Data type {data.GetType()} does not match data type of column data {reader.GetFieldType(ColumnIndexNumber)}.  Supply data with a consistent data type.");
                                errorDetected = true;
                            }
                            else
                            {
                            }
                        }
                        //else
                        //{
                        //    Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is empty. Supply valid cell data.");
                        //    errorDetected = false;
                        //}

                    }
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }

        }
    }
}






//using ExcelDataReader;
//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace ExcelDataReaderConsoleApp
//{
//    public class BasicValidation
//    {
//        private ErrorMessageLogger errorMessageLogger;
//        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();


//        public BasicValidation(ErrorMessageLogger ErrorMessageLogger)
//        {
//            errorMessageLogger = ErrorMessageLogger;

//        }
//        public Boolean NoFilesInAFolder(string[] fileEntries, string path)
//        {
//            if (fileEntries.Length == 0)
//            {
//                //errorMessageLogger.WriteToConsoleAndLogFile($"There are no files in the folder {path}, Review the folder");
//                Logger.     Error($"There are no files in the folder {path}, Review the folder");
//                //LogWriter log = new LogWriter($"There are no files in the folder {path}, Review the folder");
//                //log.LogWrite();
//                //Console.WriteLine($"There are no files in the folder {path}");
//                Console.ReadKey();
//                return true;
//            }
//            else
//            {
//                return false;
//            }
//        }

//        public bool SheetNameValidation(string fileName)
//        {
//            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
//            {
//                using (var reader = ExcelReaderFactory.CreateReader(stream))
//                {
//                    {
//                        reader.Read();
//                        {
//                            string SheetName = reader.Name;
//                            if (SheetName.Length > 128)
//                            {
//                                //errorMessageLogger.WriteToConsoleAndLogFile($" FileName: {SheetName} exceeds 128 limit, if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
//                                Logger.Error($" FileName: {SheetName} exceeds 128 limit, if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
//                                //LogWriter log = new LogWriter($" FileName: {SheetName} exceeds 128 limit, if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
//                                //log.LogWrite();
//                                //Console.WriteLine($" FileName: {SheetName} exceeds 128 limit, if the file name exceeds 128 characters, sql doesnt take file names longer than 128 characters");
//                                return true;
//                            }
//                            else
//                            {
//                                return false;
//                            }
//                        }
//                    }
//                }
//            }
//        }


//        public bool FileNameValidation(string fileName)
//        {
//            if (fileName.Length > 128)
//            {
//                //errorMessageLogger.WriteToConsoleAndLogFile($"{fileName} exceeds 128 characters, Sql doesnt take files with file name size longer than 128 characters");
//                Logger.Error($"{fileName} exceeds 128 characters, Sql doesnt take files with file name size longer than 128 characters");
//                //LogWriter log = new LogWriter($"{fileName} exceeds 128 characters, Sql doesnt take files with file name size longer than 128 characters");
//                //log.LogWrite();
//                //Console.WriteLine($"{fileName} exceeds 128 characters, Sql doesnt take files with file name size longer than 128 characters");
//                return true;
//            }
//            else
//            {
//                return false;
//            }
//        }

//        public Boolean InvalidColumnNames(string file)
//        {
//            using (FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read))
//            {
//                using (var reader = ExcelReaderFactory.CreateReader(stream))
//                {
//                    reader.Read();
//                    {
//                        int counter = 0;
//                        var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
//                        if (ColumnsNames.Count != 0 && reader.Read() == true)
//                        {
//                            foreach (var columnName in ColumnsNames)
//                            {
//                                counter += 1;
//                                if (columnName != null && columnName.ToString().Length > 128)
//                                {
//                                    //errorMessageLogger.WriteToConsoleAndLogFile($"Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name}. ColumnName: {columnName.ToString()}, if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction");
//                                    Logger.Error($"Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name}. ColumnName: {columnName.ToString()}, if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction");
//                                    //LogWriter log = new LogWriter($"Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name}. ColumnName: {columnName.ToString()}, if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction");
//                                    //log.LogWrite();
//                                    ////Console.WriteLine(i.ToString().Length);
//                                    //Console.WriteLine($" Column number {counter} is {columnName.ToString().Length} characters long in fileName: {reader.Name } if the column names length exceeds 128 characters, Sql cant upload such files because of the column names size restriction.");
//                                    //Console.WriteLine($"Column Name: {columnName.ToString()}");
//                                    //Console.ReadKey();
//                                    return true;
//                                }
//                                else if (columnName == null)
//                                {
//                                   // errorMessageLogger.WriteToConsoleAndLogFile($"Column number {counter} is missing in fileName: {reader.Name}, if the file doesnt have column names, Sql doesnt upload such files.");
//                                    Logger.Error($"Column number {counter} is missing in fileName: {reader.Name}, if the file doesnt have column names, Sql doesnt upload such files.");
//                                    //LogWriter log = new LogWriter($"Column number {counter} is missing in fileName: {reader.Name}, if the file doesnt have column names, Sql doesnt upload such files.");
//                                    //log.LogWrite();
//                                    //Console.WriteLine($" Column number {counter} is missing in fileName: {reader.Name}");
//                                    //Console.ReadKey();
//                                    return true;
//                                }
//                                else
//                                {
//                                    //Console.WriteLine(reader.Name);
//                                    //Console.WriteLine($" Column number {counter} has acceptable size in fileName: {reader.Name}, ColumnName: {columnName.ToString()}");
//                                    //Console.ReadKey();
//                                    //return false;
//                                }
//                            }
//                        }
//                        else
//                        {
//                            //errorMessageLogger.WriteToConsoleAndLogFile($" FileName: {reader.Name} is empty, if the file is empty, Sql doesnt upload such files.");
//                            Logger.Error($" FileName: {reader.Name} is empty, if the file is empty, Sql doesnt upload such files.");
//                            //LogWriter log = new LogWriter($" FileName: {reader.Name} is empty, if the file is empty, Sql doesnt upload such files.");
//                            //log.LogWrite();
//                            //Console.WriteLine($" FileName: {reader.Name} is empty");
//                            //Console.ReadKey();
//                            return true;
//                        };
//                    }
//                    return false;
//                }
//            }
//        }

//        public bool InvalidCellData(string file)
//        {
//            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
//            {
//                IExcelDataReader reader;
//                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
//                var conf = new ExcelDataSetConfiguration
//                {
//                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
//                    {
//                        UseHeaderRow = true
//                    }
//                };
//                var dataSet = reader.AsDataSet(conf);
//                var dataTable = dataSet.Tables[0];
//                var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
//                for (var ColumnIndexNumber = 0; ColumnIndexNumber < dataTable.Columns.Count; ColumnIndexNumber++)
//                {
//                    for (var RowIndexNumber = 0; RowIndexNumber < dataTable.Rows.Count; RowIndexNumber++)
//                    {
//                        //RowIndexNumber is row number starts from index number 0
//                        //ColumnIndexNumber is column number starts from index number 0
//                        var data = dataTable.Rows[RowIndexNumber][ColumnIndexNumber];
//                        if (data.GetType() == reader.GetFieldType(ColumnIndexNumber) && data.ToString().Length > 128)
//                        {
//                           // errorMessageLogger.WriteToConsoleAndLogFile($" FileName: {reader.Name} has data in a row {RowIndexNumber} column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data }, if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
//                            Logger.Error($" FileName: {reader.Name} has data in a row {RowIndexNumber} column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data }, if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
//                            //LogWriter log = new LogWriter($" FileName: {reader.Name} has data in a row {RowIndexNumber} column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data }, if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
//                            //log.LogWrite();
//                            //Console.WriteLine($" FileName: {reader.Name} has data in a row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long. The data in the cell is: { data }. if data in a cell is above 128 characters in a file, Sql doesnt upload such files.");
//                            ////Console.ReadKey();
//                            return true;
//                        }
//                        else if (data.ToString().Length != 0 & data.ToString().Length <= 128 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
//                        {
//                            //errorMessageLogger.WriteToConsoleAndLogFile($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }, if data in a cell has wrong dataType, Sql doesnt upload such files.");
//                            Logger.Error($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }, if data in a cell has wrong dataType, Sql doesnt upload such files.");
//                            //LogWriter log = new LogWriter($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }, if data in a cell has wrong dataType, Sql doesnt upload such files.");
//                            //log.LogWrite();
//                            //Console.WriteLine($" Data in a row {RowIndexNumber} column {ColumnIndexNumber} in file {reader.Name} has dataType {data.GetType()} which doesnt match {reader.GetFieldType(ColumnIndexNumber)} (column's dataType). The data in the cell is: { data }. If data in a cell has wrong dataType, Sql doesnt upload such files.");
//                            ////Console.ReadKey();
//                            return true;
//                        }
//                        else if (data.ToString().Length == 0 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
//                        {
//                            //errorMessageLogger.WriteToConsoleAndLogFile($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}. If a file has an empty cell, Sql doesnt upload such files.");
//                            Logger.Error($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}. If a file has an empty cell, Sql doesnt upload such files.");
//                            //LogWriter log = new LogWriter($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}. If a file has an empty cell, Sql doesnt upload such files.");
//                            //log.LogWrite();
//                            //Console.WriteLine($" FileName: {reader.Name} has an empty cell in row {RowIndexNumber} column {ColumnIndexNumber}. If a file has an empty cell, Sql doesnt upload such files.");
//                            ////Console.ReadKey();
//                            return false;
//                        }
//                        else if (data.ToString().Length > 128 & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
//                        {
//                            //errorMessageLogger.WriteToConsoleAndLogFile($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}, if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
//                            Logger.Error($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}, if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
//                            //LogWriter log = new LogWriter($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}, if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
//                            //log.LogWrite();
//                            //Console.WriteLine($" Data in a Row {RowIndexNumber} and column {ColumnIndexNumber} is {data.ToString().Length} characters long  and has dataType {data.GetType()}, which doesnt match column's dataType {reader.GetFieldType(ColumnIndexNumber)} in a file: {reader.Name}", "if data in a cell either is above 128 characters or has a wrong dataType or both, Sql doesnt upload such files.");
//                            ////Console.ReadKey();
//                            return true;
//                        }
//                        else
//                        {
//                            //Console.WriteLine($"{data} good size and  matches {data.GetType()}");
//                            //Console.ReadKey();
//                        }
//                    }
//                }
//                return false;
//            }
//        }

//        public string GetAddress(int column, int row)
//        {
//            column++;
//            StringBuilder stringbuilder = new StringBuilder();
//            do
//            {
//                column--;
//                stringbuilder.Insert(0, (char)('A' + (column % 26)));
//                column /= 26;

//            } while (column > 0);
//            stringbuilder.Append(row + 1);
//            return stringbuilder.ToString();
//        }
//    }
//}
