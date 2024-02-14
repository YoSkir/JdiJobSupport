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
            ShowDatabase();
        }

        private void ShowDatabase()
        {
            using (SQLiteConnection connection = dBController.GetConnection(Const.DatabaseFileName))
            {
                string sqlstr = "";
                sqlstr +=
                    "SELECT report_date AS '日期', project_code AS '專案編號', project_name AS '專案名稱', hour_spent AS '工時' FROM work_hour ";
                using(SQLiteDataAdapter adapter=new SQLiteDataAdapter(sqlstr, connection))
                {
                    DataTable dataTable=new DataTable();
                    adapter.Fill(dataTable);
                    DBDataGrid.ItemsSource = dataTable.DefaultView;
                }
            }
        }

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
