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
        private WeeklyReportPage? weeklyReportPage;
        private WorkHourReportPage? workHourReportPage;
        private DBController? dbController;

        private List<TodayReportPanel> todayPanels = new List<TodayReportPanel>();
        private List<TomorrowReportPanel> tomorrowPanels = new List<TomorrowReportPanel>();

        public MainWindow()
        {
            InitializeComponent();
            initialData();
            logLabel.Content = "輸出完成後請開啟並檢查輸出的報表。 如有未選擇專案、留空或工時填寫錯誤將不會存進資料庫，請填寫完全。" +
                "\n日報表資料有儲存功能，輸出日報表即儲存，選擇日期即可讀取當日填寫的報表。";
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            resultLabel.Content = "";
            setDate();
            writeDatabaseAndExcel();
            //重置工時表按鈕狀態
            checkWorkHourDB();
        }

        private void settingWindowButton_Click(object sender, RoutedEventArgs e)
        {
            OpenSettingWindow();
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
        /// 開啟週報表頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void weeklyReportSheet_Click(object sender, RoutedEventArgs e)
        {
            setDate();
            weeklyReportPage = new WeeklyReportPage(this);
            weeklyReportPage.Show();
        }
        /// <summary>
        /// 開啟工時表頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void workHourSheet_Click(object sender, RoutedEventArgs e)
        {
            setDate();
            dbController = new DBController(this);
            workHourReportPage = new WorkHourReportPage(this);
            workHourReportPage.Show();
        }



        /// <summary>
        /// 初始化設定
        /// </summary>
        private void initialData()
        {
            GetSourceController();
            //判斷設定異常時打開設定視窗 否則初始化日報表面板
            //改成設定部完成一樣能打開，但功能封鎖
            if (!sourceController.SourceCheck())
            {
                LockFunction();
                //OpenSettingWindow();
            }
            else
            {
                checkWorkHourDB();
                InitialPanel();
            }
        }

        private void checkWorkHourDB()
        {
            if (dbController.SelectWorkHourReport().Count == 0)
            {
                workHourPageButton.IsEnabled = false;
                workHourLabel.Content = "工時表紀錄為空";
            }
            else
            {
                workHourLabel.Content = "";
                workHourPageButton.IsEnabled = true;

            }
        }

        /// <summary>
        /// 將部分功能關閉以避免錯誤操作
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        private void LockFunction()
        {
            clearPanel();
            settingLabel.Content = "請進行設定";
            saveFileButton.IsEnabled = false;
            deleteTodayDataButton.IsEnabled = false;
            todayButton.IsEnabled = false;
            weeklyReportPageButton.IsEnabled = false;
            workHourPageButton.IsEnabled = false;
            datePicker.IsEnabled = false;
        }

        private void clearPanel()
        {
            todayPanels.Clear();
            tomorrowPanels.Clear();
            ShowPanel();
        }

        private void unlockFunction()
        {
            settingLabel.Content = "";
            saveFileButton.IsEnabled = true;
            deleteTodayDataButton.IsEnabled = true;
            todayButton.IsEnabled = true;
            weeklyReportPageButton.IsEnabled = true;
            workHourPageButton.IsEnabled = true;
            datePicker.IsEnabled = true;
        }

        /// <summary>
        /// 初始化日報表面板
        /// </summary>
        private void InitialPanel()
        {
            setDate();
            string date = defaultSetting.Default.date;
            todayPanels = dbController.GetTodayPanelList(date);
            tomorrowPanels = dbController.GetTomorrowPanelList(date);
            if (todayPanels == null || todayPanels.Count == 0)
            {
                todayPanels = new List<TodayReportPanel>();
                TodayReportPanel panel = new TodayReportPanel(this);
                todayPanels.Add(panel);
            }
            if (tomorrowPanels == null || tomorrowPanels.Count == 0)
            {
                tomorrowPanels = new List<TomorrowReportPanel>();
                TomorrowReportPanel panel = new TomorrowReportPanel(this);
                tomorrowPanels.Add(panel);
            }
            ShowPanel();
        }
        /// <summary>
        /// 檢查並返回SourceController
        /// </summary>
        /// <returns></returns>
        internal SourceController GetSourceController()
        {
            if (sourceController == null) sourceController = new SourceController(this);
            return sourceController;
        }
        internal DBController GetDBController()
        {
            if (dbController == null)
            {
                dbController = new DBController(this);
            }
            return dbController;
        }
        internal WorkHourReportPage GetWorkHourReportPage()
        {
            if (workHourReportPage == null) workHourReportPage = new WorkHourReportPage(this);
            return workHourReportPage;
        }
        /// <summary>
        /// 檢查資料庫是否有重複資料?輸出並儲存日報表:不做任何事
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void writeDatabaseAndExcel()
        {
            if (sourceController == null) GetSourceController();
            if (godModeCheckBox.IsChecked == true || sourceController.SourceCheck())
            {
                if (checkPanelInput())
                {
                    WriteDailyReport();
                }
            }
        }
        private bool checkPanelInput()
        {
            if (sourceController.CheckPanelInput(todayPanels))
                return sourceController.CheckPanelInput(tomorrowPanels);
            return false;
        }

        /// <summary>
        /// 讀取日期選擇
        /// </summary>
        private void setDate()
        {
            //如無選擇日期則默認今天
            defaultSetting.Default.date = datePicker.Text.Length > 0 ?
                datePicker.SelectedDate?.ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd");
        }
        /// <summary>
        /// 將日報表面板內容寫入檔案
        /// </summary>
        private void WriteDailyReport()
        {
            try
            {
                //做報表種類判斷
                sourceController.WritePanelToExcel(todayPanels, tomorrowPanels);
                resultLabel.Content = "儲存成功";
            }
            catch (Exception ex)
            {
                resultLabel.Content = "儲存失敗";
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 打開設定視窗
        /// </summary>
        private void OpenSettingWindow()
        {
            SettingWindow settingWindow = new SettingWindow();
            //設定視窗關閉時檢查設定失敗則鎖定功能，成功則解除鎖定
            settingWindow.Closing += SettingWindow_Closing;
            //settingWindow.Owner = this;
            settingWindow.ShowDialog();
        }
        /// <summary>
        /// 關閉視窗後檢查設定並判斷功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (sourceController == null) GetSourceController();
            if (!sourceController.SourceCheck())
            {
                LockFunction();
            }
            else
            {
                unlockFunction();
                InitialPanel();
            }
        }
        /// <summary>
        /// 重整畫面將欄位顯示於畫面上
        /// </summary>
        public void ShowPanel()
        {
            //清空面板
            todayJobPanel.Children.Clear();
            tomorrowPanelContainer.Children.Clear();
            for (int i = 0; i < todayPanels.Count; i++)
            {
                todayPanels[i].SetPanelNum(i + 1);
                todayJobPanel.Children.Add(todayPanels[i].GetPanel());
            }
            for (int i = 0; i < tomorrowPanels.Count; i++)
            {
                tomorrowPanels[i].SetPanelNum(i + 1);
                tomorrowPanelContainer.Children.Add(tomorrowPanels[i].GetPanel());
            }
        }

        internal List<TodayReportPanel> GetTodayPanels()
        {
            return todayPanels;
        }
        internal List<TomorrowReportPanel> GetTomorrowPanels()
        {
            return tomorrowPanels;
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
        /// 面板異常時重設面板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void resetPanel_Click(object sender, RoutedEventArgs e)
        {
            InitialPanel();
        }

        private void deleteTodayDataButton_Click(object sender, RoutedEventArgs e)
        {
            if (WarningBox("刪除今日(選擇的日期)的日報表、工時表紀錄?"))
            {
                string reportDate = defaultSetting.Default.date;
                dbController.DeleteOneDayData(reportDate);
                InitialPanel();
            }
        }

        private void datePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setDate();
            if (WarningBox("更換日期是否讀取當天之報表?\n(如果你報表剛打好還沒儲存，只是要換日期請按否)"))
            {
                InitialPanel();
            }
        }

        private bool WarningBox(string msg)
        {
            MessageBoxResult result = MessageBox.Show(msg, "警告", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }
            return false;
        }
    }
}