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
using NPOI.HSSF.Record;
using JDI_ReportMaker.Util.PanelComponent;

namespace JDI_ReportMaker
{
    public class DBController
    {
        private SQLiteConnection? connection;
        private MainWindow mainWindow;



        public DBController(MainWindow mainWindow)
        {
            try
            {
                connection = GetConnection(Const.DatabaseFileName);
                Connect();
                CreateTable();
                CreateDailyReportTable();
                CreateTomorrowReportTable();
                CreateWeekReportTable();
            }
            catch (Exception ex) { MessageBox.Show("資料庫創建失敗"); }

            this.mainWindow = mainWindow;
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
            SQLiteCommand cmd = new SQLiteCommand(sqlStr, connection);
            cmd.ExecuteNonQuery();
        }
        /// <summary>
        /// 創建儲存日報表的Table
        /// </summary>
        private void CreateDailyReportTable()
        {
            string sqlStr = "";
            sqlStr +=
                "CREATE TABLE IF NOT EXISTS daily_report_today " +
                "(" +
                "report_date DATETIME NOT NULL," +
                "project_name VARCHAR(100) NOT NULL, " +
                "project_title VARCHAR(100) NOT NULL, " +
                "project_describe VARCHAR(100) NOT NULL, " +
                "hour_spent TINYINT(5) NOT NULL, " +
                "jobDone BOOLEAN NOT NULL" +
                "); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlStr, connection);
            cmd.ExecuteNonQuery();
        }

        internal void InsertTodayPanel(string reportTime, string projectName, string projectTitle, string projectDescr, string hourSpend, bool jobDone)
        {
            string sqlstr =
                "INSERT INTO daily_report_today " +
                "(report_date,project_name,project_title,project_describe,hour_spent,jobDone) " +
                "VALUES " +
                $"('{reportTime}','{projectName}','{projectTitle}','{projectDescr}',{hourSpend},'{jobDone}'); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlstr, connection);
            cmd.ExecuteNonQuery();
        }
        private void DeleteTargetDateTodayPanel(string reportDate)
        {
            string paramName = "@TargetDate";
            string sqlstr =
                "DELETE FROM daily_report_today " +
               $"WHERE report_date={paramName} ";
            ExecuteDeleteCmd(sqlstr, paramName, reportDate,"日報表今日事項");
        }

        public List<TodayReportPanel> GetTodayPanelList(string reportDate)
        {
            string sqlStr = GetTodayPanelListSqlStr(reportDate);
            List<TodayReportPanel> recordList = new List<TodayReportPanel>();
            var command = new SQLiteCommand(sqlStr, connection);
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    TodayReportPanel panel = new TodayReportPanel(mainWindow);

                    panel.SetComboBoxText(reader["project_name"].ToString());
                    panel.SetTitle(reader["project_title"].ToString());
                    panel.SetDescribtion(reader["project_describe"].ToString());
                    panel.SetWorkHour(reader["hour_spent"].ToString());
                    panel.SetDone(Convert.ToBoolean(reader["jobDone"]));
                    panel.SetPanelList(recordList);
                    recordList.Add(panel);
                }
            }
            return recordList;
        }

        private string GetTodayPanelListSqlStr(string reportDate)
        {
            string sqlstr = "";
            sqlstr +=
                "SELECT project_name,project_title,project_describe,hour_spent,jobDone " +
                "FROM daily_report_today " +
               $"WHERE report_date='{reportDate}' ";
            return sqlstr;
        }

        private void CreateTomorrowReportTable()
        {
            string sqlStr = "";
            sqlStr +=
                "CREATE TABLE IF NOT EXISTS daily_report_tomorrow " +
                "(" +
                "report_date DATETIME NOT NULL," +
                "project_title VARCHAR(100) NOT NULL, " +
                "project_describe VARCHAR(100) NOT NULL " +
                "); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlStr, connection);
            cmd.ExecuteNonQuery();
        }

        internal void InsertTomorrowPanel(string reportTime,string title, string describtion)
        {
            string sqlstr =
                "INSERT INTO daily_report_tomorrow " +
                "(report_date,project_title,project_describe) " +
                "VALUES " +
                $"('{reportTime}','{title}','{describtion}'); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlstr, connection);
            cmd.ExecuteNonQuery();
        }

        private void DeleteTargetDateTomorrowPanel(string reportDate)
        {
            string paramName = "@TargetDate";
            string sqlstr =
                "DELETE FROM daily_report_tomorrow " +
               $"WHERE report_date={paramName} ";
            ExecuteDeleteCmd(sqlstr, paramName, reportDate,"日報表明日事項");
        }

        public List<TomorrowReportPanel> GetTomorrowPanelList(string reportDate)
        {
            string sqlStr = GetTomorrowPanelListSqlStr(reportDate);
            List<TomorrowReportPanel> recordList = new List<TomorrowReportPanel>();
            var command = new SQLiteCommand(sqlStr, connection);
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    TomorrowReportPanel panel = new TomorrowReportPanel(mainWindow);

                    panel.SetTitle(reader["project_title"].ToString());
                    panel.SetDescribtion(reader["project_describe"].ToString());
                    panel.SetPanelList(recordList);
                    recordList.Add(panel);
                }
            }
            return recordList;
        }

        private string GetTomorrowPanelListSqlStr(string reportDate)
        {
            string sqlstr = "";
            sqlstr +=
                "SELECT project_title,project_describe " +
                "FROM daily_report_tomorrow " +
               $"WHERE report_date='{reportDate}' ";
            return sqlstr;
        }

        /// <summary>
        /// 創建儲存周報表的Table
        /// </summary>
        private void CreateWeekReportTable()
        {
            string sqlStr = "";
            sqlStr +=
                "CREATE TABLE IF NOT EXISTS week_report " +
                "(" +
                "year tinyint(8) NOT NULL, " +
                "month tinyint(4) NOT NULL, " +
                "mon_title1 VARCHAR(100), " +
                "mon_title2 VARCHAR(100), " +
                "mon_discr1 VARCHAR(100), " +
                "mon_discr2 VARCHAR(100), " +

                "tue_title1 VARCHAR(100), " +
                "tue_title2 VARCHAR(100), " +
                "tue_discr1 VARCHAR(100), " +
                "tue_discr2 VARCHAR(100), " +

                "wed_title1 VARCHAR(100), " +
                "wed_title2 VARCHAR(100), " +
                "wed_discr1 VARCHAR(100), " +
                "wed_discr2 VARCHAR(100), " +

                "thr_title1 VARCHAR(100), " +
                "thr_title2 VARCHAR(100), " +
                "thr_discr1 VARCHAR(100), " +
                "thr_discr2 VARCHAR(100), " +

                "fri_title1 VARCHAR(100), " +
                "fri_title2 VARCHAR(100), " +
                "fri_discr1 VARCHAR(100), " +
                "fri_discr2 VARCHAR(100), " +

                "sat_title1 VARCHAR(100), " +
                "sat_title2 VARCHAR(100) " +
                "); ";
            SQLiteCommand cmd = new SQLiteCommand(sqlStr, connection);
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

        public List<String> SelectWorkHourReport()
        {
            string sqlstr =
                "SELECT project_name " +
                "FROM work_hour ";
            List<string> list = GetStrsFromDB(sqlstr);
            return list;
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

        private void DeleteTargetDateReport(string reportDate)
        {
            string paramName="@TargetDate";
            string sqlstr =
                "DELETE FROM work_hour " +
               $"WHERE report_date={paramName} ";
            ExecuteDeleteCmd(sqlstr, paramName, reportDate,"工時表紀錄 ");
        }

        public void DeleteTestData()
        {
            string paramName = "@TargetDate";
            string sqlStr =
                "DELETE FROM work_hour " +
                "WHERE report_date< @TargetDate ";
            ExecuteDeleteCmd(sqlStr, paramName, "2024-2-16","全部資料");
        }

        private void ExecuteDeleteCmd(string sqlStr,string paramName,string paramValue, string msg)
        {
            using(var command=new SQLiteCommand(sqlStr, connection))
            {
                //使用addWithValue避免資料庫注入攻擊
                command.Parameters.AddWithValue(paramName, paramValue);
                int deletedRow=command.ExecuteNonQuery();
                MessageBox.Show($"{msg} {deletedRow}項資料已移除");
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
               $"WHERE strftime('%Y-%m',report_date)='{yearMonth}' " +
               $"GROUP BY project_name ";
            return sqlStr;
        }
        /// <summary>
        /// 獲得顯示在面板上的特定月份專案工時紀錄
        /// </summary>
        /// <param name="yearMonth"></param>
        /// <returns></returns>
        public SQLiteDataAdapter SelectMonthlyData(string yearMonth)
        {
            string sqlStr=GetProjectRecordListSqlStr(yearMonth);
            return new SQLiteDataAdapter(sqlStr, connection);
        }
        /// <summary>
        /// 獲得專案紀錄的容器list
        /// </summary>
        /// <param name="yearMonth"></param>
        /// <returns></returns>
        public List<WorkHourEntity> GetProjectRecordList(string yearMonth)
        {
            string sqlStr = GetProjectRecordListSqlStr(yearMonth);
            List<WorkHourEntity> recordList= new List<WorkHourEntity>();
            var command = new SQLiteCommand(sqlStr, connection);
            using(var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    WorkHourEntity entity = new WorkHourEntity
                    {
                        projectCode = reader["專案編號"].ToString(),
                        projectName = reader["專案名稱"].ToString(),
                        recordDate = reader["日期"].ToString(),
                        hourSpent = Convert.ToInt32(reader["工時"])
                    };
                    recordList.Add(entity);
                }
            }
            return recordList;
        }
        /// <summary>
        /// 獲得專案紀錄的sql語法字串
        /// </summary>
        /// <param name="yearMonth"></param>
        /// <returns></returns>
        private string GetProjectRecordListSqlStr(string yearMonth)
        {
            string sqlstr = "";
            sqlstr +=
                "SELECT report_date AS '日期', project_code AS '專案編號', project_name AS '專案名稱', hour_spent AS '工時' " +
                "FROM work_hour " +
               $"WHERE strftime('%Y-%m',report_date)='{yearMonth}' " +
               $"ORDER BY report_date ";
            return sqlstr;
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

        public void DeleteOneDayData(string reportDate)
        {
            DeleteTargetDateReport(reportDate);
            DeleteTargetDateTodayPanel(reportDate);
            DeleteTargetDateTomorrowPanel(reportDate);
        }
    }
}
