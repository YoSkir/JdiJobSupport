using JDI_ReportMaker.Util;
using JDI_ReportMaker.Util.PanelComponent;
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
        private DBController? dBController;
        internal List<SumaryPanel> sumaryPanelList = new List<SumaryPanel>();

        public WorkHourReportPage(DBController dBController)
        {
            InitializeComponent();
            this.dBController = dBController;
            SetComboBox();
            ShowPanelAndDB(DateTime.Now.ToString("yyyy-MM"));
        }

        private void ShowPanelAndDB(string monthStr)
        {
            ShowDatabase(monthStr);
            SumaryPanelInitial(monthStr);
        }
        private void SumaryPanelInitial(string yearMonth)
        {
            SetUpPanelList(yearMonth);
            ShowSumaryPanel();
        }
        private void ShowSumaryPanel()
        {
            SumaryPanelContainer.Children.Clear();
            foreach(SumaryPanel panel in sumaryPanelList)
            {
                SumaryPanelContainer.Children.Add(panel.GetPanel());
            }
        }
        private void SetUpPanelList(string yearMonth)
        {
            string sqlStr = dBController.SelectMonthlyHourSpentByProjectName(yearMonth);
            List<WorkHourComponent> componentList = dBController.GetProjectsHourSpent(sqlStr);
            foreach (WorkHourComponent component in componentList)
            {
                SumaryPanel panel = new SumaryPanel(this);
                panel.SetPanelValue(component.project_code, component.project_name, component.hour_spent.ToString());
                panel.AddPanel();
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
            if(yearMonth.Length > 0)
            ShowPanelAndDB(yearMonth);
        }
        private string ComboBoxStrToShowDBStr(string comboBoxStr)
        {
            return comboBoxStr.Length > 0?comboBoxStr.Replace('年','-').Substring(0,7):"";
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

        private void deleteTestData_Click(object sender, RoutedEventArgs e)
        {
            dBController.DeleteTestData();
        }
    }
}
