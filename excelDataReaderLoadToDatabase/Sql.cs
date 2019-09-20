using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
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

        public Sql(String sheetName, List<dynamic> columnNames, List<dynamic> DataTypes, DataTable dataTable)
        {
            SheetName = sheetName;
            ColumnNames = columnNames;
            ColumnsDataTypes = DataTypes;
            DataTable = dataTable;
        }


        public void CreateTable()
        {
            string server = System.Configuration.ConfigurationManager.AppSettings["server"];
            string database = ConfigurationManager.AppSettings["database"].ToString();
            Console.WriteLine(server);
            Console.WriteLine(database);
            //Connect to the local, default instance of SQL Server.   
            srv = new Server(server);
            Console.WriteLine(srv);
            //Reference the AdventureWorks2012 database.   
            db = srv.Databases[database];
            //Define a Table object variable by supplying the parent database and table name in the constructor.  
            tb = new Table(db, SheetName);
            DropTableIfExists();
            CreateColumns();
            tb.Create();
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

        public void DropTableIfExists()
        {
            bool tableExists = db.Tables.Contains(SheetName);
            if (tableExists)
            {
                db.Tables[SheetName].Drop();
            }
        }

        public void CreateColumns()
        {
            for (dynamic i = 0; i < Math.Min(ColumnNames.Count, ColumnsDataTypes.Count); i++)
            {
                Column col;
                col = new Column(tb, ColumnNames[i], ColumnsDataTypes[i]);
                tb.Columns.Add(col);
            }
        }

        public void CreateRows()
        {
            // String sqlConnectionString = "Data Source = localhost\\ADM; Initial Catalog =fileToUpload; Integrated Security = SSPI;";
            string sqlConnectionString = ConfigurationManager.ConnectionStrings["MyKey"].ConnectionString;
            Console.WriteLine(sqlConnectionString);
            Console.ReadKey();
            using (var bulkCopy = new SqlBulkCopy(sqlConnectionString))
            {
                bulkCopy.DestinationTableName = tb.Name.ToString();
                bulkCopy.WriteToServer(DataTable);
            }
            //Console.WriteLine($"Copying data to the table {DataTable} in database.");
            //Console.ReadKey();
            //Console.WriteLine("database Updted.");
            //Console.ReadKey();
        }
    }
}
