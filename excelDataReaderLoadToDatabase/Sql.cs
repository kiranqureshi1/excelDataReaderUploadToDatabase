using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelDataReaderConsoleApp
{
    class Sql
    {
        public object MessageBox { get; private set; }
        private Database db;
        private Table tb;
        private Server srv;
        private readonly String SheetName;
        private List<dynamic> ColumnNames;
        private List<dynamic> ColumnsDataTypes;
        private readonly DataTable DataTable;
        private static readonly string server;

        //public Sql()
        //{
        //}
        static Sql()
        {
            server = System.Configuration.ConfigurationManager.AppSettings["server"];
        }

        public Sql(String sheetName, DataTable dataTable, List<dynamic> columnNames, List<dynamic> DataTypes)
        {
            SheetName = sheetName;
            DataTable = dataTable;
            ColumnNames = columnNames;
            ColumnsDataTypes = DataTypes;
        }

        public void CreateDbGet_Table_Columns_Rows_DataTypes(string database)
        {
            //string server = System.Configuration.ConfigurationManager.AppSettings["server"];
            //string database = ConfigurationManager.AppSettings["database"].ToString();
            srv = new Server(server);
            DropDatabaseIfExists(database);
            db = new Database(srv, database);
            db.Create();
            ConvertExcelDataTypesToSql();
            CreateTable();
            // CreateColumns();
            CreateRows(database);
        }


        public Database UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(string database)
        {
            string server = System.Configuration.ConfigurationManager.AppSettings["server"];
            //string database = ConfigurationManager.AppSettings["database"].ToString();
            srv = new Server(server);
            db = srv.Databases[database];
            ConvertExcelDataTypesToSql();
            CreateTable();
            CreateRows(database);
            return db;
        }


        public Table CreateTable()
        {
            // string database = ConfigurationManager.AppSettings["database"].ToString();
            // //string server = System.Configuration.ConfigurationManager.AppSettings["server"];

            //srv = new Server(server);
            // db = srv.Databases[database];
            // //tb = new Table(db, SheetName);
            tb = new Table(db, SheetName);
            DropTableIfExists();
            CreateColumns();
            tb.Create();
            return tb;
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

        public void DropDatabaseIfExists(string database)
        {
            Boolean databaseExists = srv.Databases.Contains(database);
            if (databaseExists)
            {
                srv.Databases[database].Drop();
            }
        }

        public void CreateColumns()
        {
            for (dynamic indexNumber = 0; indexNumber < Math.Min(ColumnNames.Count, ColumnsDataTypes.Count); indexNumber++)
            {
                Column column;
                column = new Column(tb, ColumnNames[indexNumber], ColumnsDataTypes[indexNumber])
                {
                    Collation = "Latin1_General_CI_AS",
                    Nullable = true
                };
                tb.Columns.Add(column);
            }
        }

        public void CreateRows(string database)
        {
            // string database = ConfigurationManager.AppSettings["database"].ToString();
            // SqlConnectionStringBuilder sqlConnectionString = new SqlConnectionStringBuilder();
            // sqlConnectionString.DataSource = @"localhost\ADM";
            // sqlConnectionString.InitialCatalog = database;
            //sqlConnectionString.IntegratedSecurity = true;
            // MessageBox.Show(connectionString.ConnectionString);
            string sqlConnectionString = ConfigurationManager.ConnectionStrings["MyKey"].ConnectionString;
            sqlConnectionString = string.Format(sqlConnectionString, server, database);
            using (var bulkCopy = new SqlBulkCopy(sqlConnectionString))
            {
                bulkCopy.DestinationTableName = tb.Name.ToString();
                bulkCopy.WriteToServer(DataTable);
            }
        }
    }
}
