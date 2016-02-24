using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace DBConvertToExcel
{
    class SqliteHelper
    {
        public SqliteHelper()
        {

        }

        public static SQLiteConnection GetConnection(string dbFilePath)
        {
            String connStr = "Data Source=" + dbFilePath;
            SQLiteConnection conn = new SQLiteConnection(connStr);
            return conn;
        }

        public static int ExecuteSql(string sql,string dbPath)
        {
            using (SQLiteConnection conn=GetConnection(dbPath))
            {
                var cmd = new SQLiteCommand(sql, conn);
                return cmd.ExecuteNonQuery();
            }
        }

        public static DataSet ExcelDataSet(string sql,string dbPath)
        {
            using (SQLiteConnection conn=GetConnection(dbPath))
            {
                var cmd = new SQLiteCommand(sql, conn);
                SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
                
            }
           
        }
    }
    
}
