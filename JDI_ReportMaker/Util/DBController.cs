using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows;
using JDI_ReportMaker.Util;
using NPOI.Util.Collections;

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
                Connect();
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
        }

        public void InsertWorkHour(string recordDate,string projectCode,string projectName,string hourSpent)
        {
            string sqlstr = 
                "INSERT INTO work_hour " +
                "(report_date,project_code,project_name,hour_spent) " +
                "VALUES " +
                $"('{recordDate}','{projectCode}','{projectName}',{hourSpent}); ";
            SQLiteCommand cmd=new SQLiteCommand(sqlstr,connection);
            cmd.ExecuteNonQuery();
        }

        public List<string> SelectExistDateData(string reportDate)
        {
            string sqlstr =
                "SELECT project_name " +
                "FROM work_hour " +
               $"WHERE report_date='{reportDate}' ";
            List<string> list=GetStrsFromDB(sqlstr);
            return list;
        }

        public void DeleteTargetDateReport(string reportDate)
        {
            string paramName="@TargetDate";
            string sqlstr =
                "DELETE FROM work_hour " +
               $"WHERE report_date={paramName} ";
            ExecuteDeleteCmd(sqlstr, paramName, reportDate);
        }

        public void DeleteTestData()
        {
            string paramName = "@TargetDate";
            string sqlStr =
                "DELETE FROM work_hour " +
                "WHERE report_date< @TargetDate ";
            ExecuteDeleteCmd(sqlStr, paramName, "2024-2-16");
        }

        private void ExecuteDeleteCmd(string sqlStr,string paramName,string paramValue)
        {
            using(var command=new SQLiteCommand(sqlStr, connection))
            {
                //使用addWithValue避免資料庫注入攻擊
                command.Parameters.AddWithValue(paramName, paramValue);
                int deletedRow=command.ExecuteNonQuery();
                MessageBox.Show($"{deletedRow}項資料已移除");
            }
        }

        public List<DateTime> SelectReportDate()
        {
            string sqlStr =
                "SELECT Distinct report_date " +
                "FROM work_hour " +
                "ORDER BY report_date ";
            return GetDateListFromDB(sqlStr);
        }


        public List<string> SelectMonthlyProjectName(string yearMonth)
        {
            string sqlStr =
                "SELECT project_name " +
                "FORW work_hour " +
               $"WHERE strftime(%Y-%m,report_date)='{yearMonth}' " +
                "GROUP BY project_name ";
            return GetStrsFromDB(sqlStr);
        }

        public string SelectMonthlyHourSpentByProjectName(string yearMonth)
        {
            string sqlStr =
                "SELECT project_code,project_name,SUM(hour_spent) AS hour_spent " +
                "FROM work_hour " +
               $"WHERE strftime('%Y-%m',report_date='{yearMonth}' ) " +
               $"GROUP BY project_name ";
            return sqlStr;
        }

        public SQLiteDataAdapter SelectMonthlyData(string yearMonth)
        {
            string sqlstr = "";
            sqlstr +=
                "SELECT report_date AS '日期', project_code AS '專案編號', project_name AS '專案名稱', hour_spent AS '工時' " +
                "FROM work_hour " +
               $"WHERE strftime('%Y-%m',report_date)='{yearMonth}' ";
            return new SQLiteDataAdapter(sqlstr, connection);
        }

        public List<WorkHourEntity> GetProjectsHourSpent(string sqlStr)
        {
            List<WorkHourEntity> workHourComponents = new List<WorkHourEntity>();
            var command = new SQLiteCommand(sqlStr,connection);
            using(var reader=command.ExecuteReader())
            {
                while(reader.Read())
                {
                    WorkHourEntity component = new WorkHourEntity
                    {
                        projectCode = reader["project_code"].ToString(),
                        projectName = reader["project_name"].ToString(),
                        hourSpent = Convert.ToInt32(reader["hour_spent"])
                    };
                    workHourComponents.Add(component);
                }
            }
            return workHourComponents;
        }

        /// <summary>
        /// 將datetime sql轉成List datetime
        /// </summary>
        /// <param name="selectDatetimeSqlStr"></param>
        /// <returns></returns>
        private List<DateTime> GetDateListFromDB(string selectDatetimeSqlStr)
        {
            List<DateTime> result = new List<DateTime>();
            var command = new SQLiteCommand(selectDatetimeSqlStr, connection);
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var dateStr = reader.GetString(0);
                    DateTime date;
                    if (DateTime.TryParse(dateStr, out date))
                    {
                        result.Add(date);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 將string sql轉成List str
        /// </summary>
        /// <returns></returns>
        private List<string> GetStrsFromDB(string selectStrSqlStr)
        {
            List<string> result = new List<string>();
            var command = new SQLiteCommand(selectStrSqlStr, connection);
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var str = reader.GetString(0);
                    result.Add(str);
                }
            }
            return result;
        }
    }
}
