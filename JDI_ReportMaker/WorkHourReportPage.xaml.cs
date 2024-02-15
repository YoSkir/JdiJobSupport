using JDI_ReportMaker.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace JDI_ReportMaker
{
    /// <summary>
    /// WorkHourReportPage.xaml 的互動邏輯
    /// </summary>
    public partial class WorkHourReportPage : Window
    {
        private DBController dBController;

        public WorkHourReportPage(DBController dBController)
        {
            InitializeComponent();
            this.dBController = dBController;
            SetComboBox();
            ShowPanelAndDB(DateTime.Now);

        }

        private void ShowPanelAndDB(DateTime month)
        {
            string monthStr = month.ToString("yyyy-MM");
            ShowDatabase(monthStr);
            //ShowSumaryPanel(monthStr);
        }

        private void ShowSumaryPanel(string yearMonth)
        {
            using (SQLiteConnection connection = dBController.GetConnection(Const.DatabaseFileName))
            {
                string sqlstr = "";
                sqlstr +=
                    "SELECT report_date AS '日期', project_code AS '專案編號', project_name AS '專案名稱', hour_spent AS '工時' " +
                    "FROM work_hour " +
                   $"WHERE strftime('%Y-%m',report_date)='{yearMonth}' ";
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(sqlstr, connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    DBDataGrid.ItemsSource = dataTable.DefaultView;
                }
            }
        }

        /// <summary>
        /// combo box的設定
        /// </summary>
        private void SetComboBox()
        {
            SetComboBoxList(GetMonthListFromDB());
            SetComboBoxEvent();
        }
        //事件設定
        private void SetComboBoxEvent()
        {
            choseMonthCBox.DropDownClosed += ChoseMonthCBox_DropDownClosed;
        }
        //選單關閉時的事件
        private void ChoseMonthCBox_DropDownClosed(object? sender, EventArgs e)
        {
            string yearMonth = ComboBoxStrToShowDBStr(choseMonthCBox.Text);
            ShowDatabase(yearMonth);
        }
        private string ComboBoxStrToShowDBStr(string comboBoxStr)
        {
            return comboBoxStr.Replace('年','-').Substring(0,7);
        }
        //設定下拉選單
        private void SetComboBoxList(List<string> list)
        {
            choseMonthCBox.ItemsSource=list;
        }
        //獲得下拉選單內容
        private List<string> GetMonthListFromDB()
        {
            List<DateTime> dates = dBController.SelectReportDate();
            List<string> months=new List<string>();
            foreach (DateTime date in dates)
            {
                string monthStr = date.ToString("yyyy年MM月");
                if (!months.Contains(monthStr))
                {
                    months.Add(monthStr);
                }
            }
            return months;
        }
        /// <summary>
        /// 於面板上顯示指定月份的資料
        /// </summary>
        private void ShowDatabase(string yearMonth)
        {
            SQLiteDataAdapter adapter = dBController.SelectMonthlyData(yearMonth);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            DBDataGrid.ItemsSource = dataTable.DefaultView;
        }
        /// <summary>
        /// UI日期格式的設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DBDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if(e.PropertyType == typeof(DateTime))
            {
                var column = e.Column as DataGridTextColumn;
                if(column!=null && e.Column.Header.ToString()=="日期")
                {
                    column.Binding.StringFormat = "MM月dd日";
                }
            }
        }
    }
}
