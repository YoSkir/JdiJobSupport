using MahApps.Metro.Controls;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming.Properties;
using NPOI.Util.Collections;
using JDI_ReportMaker.ExcelWriter;
using JDI_ReportMaker.Util;
using NPOI.Util;
using JDI_ReportMaker.Util.PanelComponent;

namespace JDI_ReportMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private SourceController? sourceController;
        private List<TodayReportPanel>? todayPanels;
        private List<TomorrowReportPanel>? tomorrowPanels;
        private WeeklyReportPage? weeklyReportPage;
        private WorkHourReportPage? workHourReportPage;
        private DBController? dbController;

        public MainWindow()
        {
            InitializeComponent();
            initialData();
        }
        /// <summary>
        /// 初始化設定
        /// </summary>
        private void initialData()
        {
            GetSourceController();
            //判斷設定異常時打開設定視窗 否則初始化日報表面板
            if (!sourceController.SourceCheck())
            {
                OpenSettingWindow();
            }
            else
            {
                InitialPanel();
            }
        }
        /// <summary>
        /// 初始化日報表面板
        /// </summary>
        private void InitialPanel()
        {
            if (todayPanels == null)
            {
                todayPanels = new List<TodayReportPanel>();
                AddTodayPanel();
            }
            if (tomorrowPanels == null)
            {
                tomorrowPanels = new List<TomorrowReportPanel>();
                AddTomorrowPanel();
            }
            if (todayPanels.Count == 0)
            {
                AddTodayPanel();
            }
            if (tomorrowPanels.Count == 0)
            {
                AddTomorrowPanel();
            }
            Show();
        }

        internal SourceController GetSourceController()
        {
            if (dbController == null) { dbController=new DBController(); }
            if (sourceController == null) sourceController = new SourceController(dbController);
            return sourceController;
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            resultLabel.Content = "";
            //如無選擇日期則默認今天
            defaultSetting.Default.date=datePicker.Text.Length>0?
                datePicker.SelectedDate?.ToString("yyyy-MM-dd"): DateTime.Now.ToString("yyyy-MM-dd");
            //此布林用於警告使用者輸入有異常是否繼續
            bool inputOK = false;
            if (sourceController == null) GetSourceController();
            if (godModeCheckBox.IsChecked == true|| sourceController.SourceCheck())
            {
                inputOK = sourceController.CheckPanelInput(todayPanels);
                if(inputOK)
                inputOK = sourceController.CheckPanelInput(tomorrowPanels);
                if (inputOK)
                {
                    WriteDailyReport();
                }
            }
            else
            {
                logLabel.Content = "請確認檔案路徑、員工資料、檔案是否正常";
            }
        }
        /// <summary>
        /// 將日報表面板內容寫入檔案
        /// </summary>
        private void WriteDailyReport()
        {
            if (sourceController == null) GetSourceController();
            try
            {
                //做報表種類判斷
                sourceController.WritePanelToExcel(todayPanels, tomorrowPanels);
                resultLabel.Content = "儲存成功";
                logLabel.Content = "";
            }
            catch (Exception ex)
            {
                resultLabel.Content = "儲存失敗";
                logLabel.Content = ex.Message;
            }
        }

        private void settingWindowButton_Click(object sender, RoutedEventArgs e)
        {
            OpenSettingWindow();
        }
        /// <summary>
        /// 打開設定視窗
        /// </summary>
        private void OpenSettingWindow()
        {
            SettingWindow settingWindow = new SettingWindow();
            //設定需要檔案與設定正常才能關閉
            settingWindow.Closing += SettingWindow_Closing;
            //settingWindow.Owner = this;
            settingWindow.ShowDialog();
        }
        /// <summary>
        /// 將日期重置今天
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void todayButton_Click(object sender, RoutedEventArgs e)
        {
            datePicker.Text = "";
        }
        /// <summary>
        /// 重整畫面將欄位顯示於畫面上
        /// </summary>
        public void ShowPanel()
        {
            //清空面板
            todayJobPanel.Children.Clear();
            tomorrowPanelContainer.Children.Clear();
            if (todayPanels == null)
                todayPanels = new List<TodayReportPanel>();
            for (int i = 0; i < todayPanels.Count; i++)
            {
                todayPanels[i].SetPanelNum(i + 1);
                todayJobPanel.Children.Add(todayPanels[i].GetPanel());
            }
            if (tomorrowPanels == null)
                tomorrowPanels = new List<TomorrowReportPanel>();
            for (int i = 0; i < tomorrowPanels.Count; i++)
            {
                tomorrowPanels[i].SetPanelNum(i + 1);
                tomorrowPanelContainer.Children.Add(tomorrowPanels[i].GetPanel());
            }
        }
        /// <summary>
        /// 新增面板事項欄位
        /// </summary>
        internal void AddTodayPanel()
        {
            if(todayPanels != null && todayPanels.Count < 7)
            {
                TodayReportPanel panel = new TodayReportPanel(this);
                todayPanels.Add(panel);
                ShowPanel();
            }
        }
        internal void AddTomorrowPanel()
        {
            if(tomorrowPanels != null&&tomorrowPanels.Count < 5)
            {
                TomorrowReportPanel panel = new TomorrowReportPanel(this);
                tomorrowPanels.Add(panel);
                ShowPanel();
            }
        }
        /// <summary>
        /// 刪除面板特定欄位
        /// </summary>
        /// <param name="target">要刪除的欄位，由欄位本身回傳</param>
        internal void RemovePanel(TodayReportPanel target)
        {
            if(todayPanels!=null && todayPanels.Count > 1)
            {
                todayPanels.Remove(target);
                ShowPanel();
            }
        }
        internal void RemovePanel(TomorrowReportPanel target)
        {
            if (tomorrowPanels != null && tomorrowPanels.Count > 1)
            {
                tomorrowPanels.Remove(target);
                ShowPanel();
            }
        }
        /// <summary>
        /// 測試用 清空設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cleanDefaultButton_Click(object sender, RoutedEventArgs e)
        {
            defaultSetting.Default.Reset();
        }
        /// <summary>
        /// 設定未完成時無法關閉設定視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (sourceController == null) GetSourceController();
            if (!sourceController.SourceCheck())
            {
                e.Cancel = true;
            }
            else
            {
                InitialPanel();
            }
        }
        /// <summary>
        /// 面板異常時重設面板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void resetPanel_Click(object sender, RoutedEventArgs e)
        {
            InitialPanel();
        }

        private void weeklyReportSheet_Click(object sender, RoutedEventArgs e)
        {
            defaultSetting.Default.date = datePicker.Text.Length > 0 ?
                                          datePicker.SelectedDate?.ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd");
            weeklyReportPage = new WeeklyReportPage(this);
            weeklyReportPage.Show();
        }

        private void workHourSheet_Click(object sender, RoutedEventArgs e)
        {
            defaultSetting.Default.date = datePicker.Text.Length > 0 ?
                              datePicker.SelectedDate?.ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd");
            dbController=new DBController();
            workHourReportPage=new WorkHourReportPage(dbController);
            workHourReportPage.Show();
        }
    }
}