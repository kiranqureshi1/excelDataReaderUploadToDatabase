using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReaderConsoleApp
{
    class Sql
    {
        public object MessageBox { get; private set; }
        public static Database db;
        public static Table tb;
        public static Server srv;
        public static String SheetName;
        public static List<dynamic> ColumnNames;
        public static List<dynamic> ColumnsDataTypes;
        public static DataTable DataTable;
        public static string path = "C:\\Temp\\source";
        public static string search = "*.xlsx";

        //public Sql()
        //{
        //}

        public Sql(String sheetName, List<dynamic> columnNames, List<dynamic> DataTypes, DataTable dataTable)
        {
            SheetName = sheetName;
            ColumnNames = columnNames;
            ColumnsDataTypes = DataTypes;
            DataTable = dataTable;
        }


        public void CreateTable()
        {
            string[] fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);
            foreach (string fileName in fileEntries)
            {
                //Connect to the local, default instance of SQL Server.   
                srv = new Server(".\\");
                Console.WriteLine(srv.Name);
                //Reference the AdventureWorks2012 database.   
                db = srv.Databases["fileToUpload"];
                Console.WriteLine(db.Name);
                Console.ReadKey();
                //Define a Table object variable by supplying the parent database and table name in the constructor.  
                tb = new Table(db, SheetName);
                DropTableIfExists();
                CreateColumns();
                //CreateRows();
                tb.Create();
            }

        }

        public void ConvertExcelDataTypesToSql()

        {
            for (dynamic i = 0; i < ColumnsDataTypes.Count; i++)
            {
                if (ColumnsDataTypes[i] == typeof(String))
                {
                    ColumnsDataTypes[i] = DataType.NVarChar(255);
                }
                else
                {
                    ColumnsDataTypes[i] = DataType.Float;
                }

            }
        }

        
    }
}
