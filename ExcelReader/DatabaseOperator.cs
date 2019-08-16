using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelReader
{
    public static class DatabaseOperator
    {
        private const string ConnectionString = "Server=.\\SQLEXPRESS;Database=ExcelReader;Trusted_Connection=True;MultipleActiveResultSets=true";
      private const  string Excel03ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR={1}'";
      private const  string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
      private const  string CSVConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='text;HDR=YES;FMT=Delimited'";
       private const string TableName = "Product";
        public static SqlConnection GetDbContext()
        {
            try
            {
                SqlConnection connection = new SqlConnection(ConnectionString);

                connection.Open();

                return connection;

            }
            catch (Exception es)

            {
                MessageBox.Show(es.Message);
                return null;
            }
        }

        public static void DisposeDbConnection(SqlConnection dbContext)
        {
            dbContext.Close();
        }

        public static ICollection<string> GetColumnNames()
        {
            var dbContext = GetDbContext();
            string[] restrictions = new string[4] { null, null, "Product", null };

            var columnList = dbContext.GetSchema("Columns", restrictions).AsEnumerable().Select(s => s.Field<String>("Column_Name")).ToList();
            return columnList;
        }

        public static void ParsFile(string fullFilePath, SqlConnection dbContext)
        {
            string extension = Path.GetExtension(fullFilePath);
            string conString = string.Empty;
            switch (extension)
            {
                case ".xls":
                   ParseExcel(fullFilePath, dbContext, string.Format(Excel03ConString, fullFilePath, "YES"));
                    break;
                case ".xlsx":
                    ParseExcel(fullFilePath, dbContext,string.Format(Excel07ConString, fullFilePath, "YES"));
                    break;
                default:
                    ParseCsv(fullFilePath);  // conString = string.Format(CSVConString, fullFilePath, "NO");
                    break;
            }
        }

        public static void ParseExcel(string fullFilePath, SqlConnection dbContext, string oleDbConnectionString)
        {
            try
            {

                OleDbConnection oledbconn = new OleDbConnection(oleDbConnectionString);

                oledbconn.Open();
                var schema = oledbconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                var sheetName = schema.Rows[0]["Table_Name"].ToString();
                ;

                var dbColumnNameList = GetColumnNames();
                var fileColumnNameList = new List<string>();
                try
                {
                    var columns = oledbconn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null);
                  
                    if (columns != null)
                    {
                        fileColumnNameList.AddRange(from DataRow column in columns.Rows select column["Column_name"].ToString());
                    }

                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                }

                var mappingColumnNames = new Dictionary<string, string>();
                foreach (var name in dbColumnNameList)
                {
                    var mapp = fileColumnNameList.FirstOrDefault(f => f == name);
                    if (mapp != null)
                    {
                        mappingColumnNames.Add(name, mapp);
                        fileColumnNameList.Remove(name);
                    }

                }

                SqlBulkCopy bulkcopy = new SqlBulkCopy(ConnectionString);

                foreach (var item in mappingColumnNames)
                {
                    bulkcopy.ColumnMappings.Add(item.Key, item.Value);
                }

                string exceldataquery = $"select {string.Join(",", mappingColumnNames.Values)} from [{sheetName}]";
                OleDbCommand oledbcmd = new OleDbCommand(exceldataquery, oledbconn);
                OleDbDataReader dr = oledbcmd.ExecuteReader();

                bulkcopy.DestinationTableName = TableName;
                bulkcopy.WriteToServer(dr);
                bulkcopy.Close();
                oledbconn.Close();
            }
            catch (Exception )
            {
                MessageBox.Show("Invalid Operation!!!");
            }
        }
        public static void ParseCsv(string filename)
        {
            var lines = System.IO.File.ReadAllLines(filename);
            if (lines.Count() == 0) return;
            var columns = lines[0].Split(',');
            var table = new DataTable();
            var fileColumnNameList = new List<string>();
            foreach (var c in columns)
            {
                table.Columns.Add(c);
                fileColumnNameList.Add(c);
            }


            for (int i = 1; i < lines.Count() - 1; i++)
            {
                table.Rows.Add(lines[i].Split(','));
            }
            var dbColumnNameList = GetColumnNames();
            var mappingColumnNames = new Dictionary<string, string>();
            foreach (var name in dbColumnNameList)
            {
                var mapp = fileColumnNameList.FirstOrDefault(f => f == name);
                if (mapp != null)
                {
                    mappingColumnNames.Add(name, mapp);
                    fileColumnNameList.Remove(name);
                }

            }

            mappingColumnNames.Keys.ToList().ForEach(key => dbColumnNameList.Remove(key));

           
            DataTable dtInsertRows = table;

            using (SqlBulkCopy sbc = new SqlBulkCopy(ConnectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                sbc.DestinationTableName = TableName;

                if (dbColumnNameList.Count != 0)
                {
                    foreach (var name in dbColumnNameList)
                    {
                        if (fileColumnNameList.Contains(name))
                        {
                            var mappingFileColumnName = fileColumnNameList.First(columnName => columnName.Contains(name));
                            mappingColumnNames.Add(name, mappingFileColumnName);
                        }
                    }
                   
                }
             
                sbc.BatchSize = table.Rows.Count;

                foreach (var item in mappingColumnNames)
                {
                    sbc.ColumnMappings.Add(item.Key, item.Value);
                }
                sbc.WriteToServer(dtInsertRows);
            }
        }
    }
}