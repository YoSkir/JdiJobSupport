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
        private List<TodayReportPanel>? colums;

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
            savePathTextBox.Text = defaultSetting.Default.target_path_d;
            sourceController = new SourceController();
            colums = new List<TodayReportPanel>();
            AddPanel();
            ShowPanel();
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            resultLabel.Content = "";
            defaultSetting.Default.target_path_d = savePathTextBox.Text;
            defaultSetting.Default.Save();
            defaultSetting.Default.date=datePicker.Text.Length>0?
                datePicker.SelectedDate?.ToString("yyyy-MM-dd"): DateTime.Now.ToString("yyyy-MM-dd");
            if (sourceController != null && (godModeCheckBox.IsChecked == true|| sourceController.SourceCheck()))
            {
                try
                {
                    //做報表種類判斷
                    sourceController.ExcecuteFile(colums);
                    resultLabel.Content = "儲存成功";
                    logLabel.Content = "";
                }catch (Exception ex)
                {
                    resultLabel.Content = "儲存失敗";
                    logLabel.Content = ex.Message;
                }
            }
            else
            {
                logLabel.Content = "請確認檔案路徑、員工資料、檔案是否正常";
            }
        }
        private void settingWindowButton_Click(object sender, RoutedEventArgs e)
        {
            SettingWindow settingWindow = new SettingWindow();
            settingWindow.Owner = this;
            settingWindow.ShowDialog();
        }

        private void selectLocateButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFolderDialog openFolderDialog = new OpenFolderDialog() {
                Title = "請選擇目標位置",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                Multiselect = false
            } ;
            if (openFolderDialog.ShowDialog() == true)
            {
                savePathTextBox.Text = openFolderDialog.FolderName;
            }
        }

        private void todayButton_Click(object sender, RoutedEventArgs e)
        {
            datePicker.Text = "";
        }
        /// <summary>
        /// 重整畫面將欄位顯示於畫面上
        /// </summary>
        public void ShowPanel()
        {
            todayJobPanel.Children.Clear();
            if (colums != null)
                for (int i = 0; i < colums.Count; i++)
                {
                    colums[i].SetPanelNum(i + 1);
                    todayJobPanel.Children.Add(colums[i].GetPanel());
                }
        }
        /// <summary>
        /// 新增欄位
        /// </summary>
        internal void AddPanel()
        {
            if(colums != null && colums.Count < 7)
            {
                TodayReportPanel panel = new TodayReportPanel(this);
                colums.Add(panel);
                ShowPanel();
            }
        }
        /// <summary>
        /// 刪除欄位
        /// </summary>
        /// <param name="target">要刪除的欄位，由欄位本身回傳</param>
        internal void RemovePanel(TodayReportPanel target)
        {
            if(colums!=null && colums.Count > 1)
            {
                colums.Remove(target);
                ShowPanel();
            }
        }
    }
}