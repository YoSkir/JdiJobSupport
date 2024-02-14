using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows;

namespace JDI_ReportMaker
{
    internal class DBController
    {
        private SQLiteConnection connection;



        public DBController()
        {
            try
            {
                string connectionStr = "Data Source=JJSDatabase.sqlite;Version=3;";
                connection = new SQLiteConnection(connectionStr);
                CreateTable();
            }catch (Exception ex) { MessageBox.Show("資料庫創建失敗"); }

        }

        private void Connect()
        {
            if (connection.State != System.Data.ConnectionState.Open)
            {
                connection.Open();
            }
        }

        private void Disconnect()
        {
            if(connection.State != System.Data.ConnectionState.Closed)
            {
                connection.Close();
            }
        } 

        private void CreateTable()
        {
            Connect();
            string sqlStr = "";
            sqlStr += 
                "CREATE TABLE IF NOT EXISTS work_hour " +
                "(" +
                "report_date DATETIME NOT NULL," +
                "project_code VARCHAR(15), " +
                "project_name VARCHAR(100), " +
                "hour_spent TINYINT(5) " +
                "); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlStr,connection);
            cmd.ExecuteNonQuery();
            Disconnect();
        }
        
    }
}
