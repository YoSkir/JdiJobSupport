using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows;
using JDI_ReportMaker.Util;

namespace JDI_ReportMaker
{
    public class DBController
    {
        private SQLiteConnection? connection;



        public DBController()
        {
            try
            {
                connection = GetConnection(Const.DatabaseFileName);
                CreateTable();
            }catch (Exception ex) { MessageBox.Show("資料庫創建失敗"); }

        }

        public SQLiteConnection GetConnection(string dbName)
        {
            string connectionStr = $"Data Source={dbName};Version=3;";
            connection = new SQLiteConnection(connectionStr);
            return connection;
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

        public void InsertWorkHour(string recordDate,string projectCode,string projectName,string hourSpent)
        {
            Connect();
            string sqlstr = 
                "INSERT INTO work_hour " +
                "(report_date,project_code,project_name,hour_spent) " +
                "VALUES " +
                $"('{recordDate}','{projectCode}','{projectName}',{hourSpent}); ";
            SQLiteCommand cmd=new SQLiteCommand(sqlstr,connection);
            cmd.ExecuteNonQuery();
            Disconnect();
        }
        
    }
}
