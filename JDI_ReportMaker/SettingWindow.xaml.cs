using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
    /// PopupWindow.xaml 的互動邏輯
    /// </summary>
    public partial class SettingWindow : Window
    {
        public SettingWindow()
        {
            InitializeComponent();
            initialData();
        }
        /// <summary>
        /// 初始化資料
        /// </summary>
        private void initialData()
        {
            dailyReportTbox.Text = defaultSetting.Default.source_path_d;
            weeklyReportTbox.Text = defaultSetting.Default.source_path_w;
            workHourReportTbox.Text=defaultSetting.Default.source_path_h;
            staffNameTBox.Text = defaultSetting.Default.staff_name;
            departmentCbox.Text = defaultSetting.Default.department;
            savePathTextBox.Text = defaultSetting.Default.target_path_d;
        }
        /// <summary>
        /// 將設定寫入預設
        /// </summary>
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                defaultSetting.Default.source_path_d = dailyReportTbox.Text;
                defaultSetting.Default.source_path_w = weeklyReportTbox.Text;
                defaultSetting.Default.source_path_h = workHourReportTbox.Text;
                defaultSetting.Default.staff_name = staffNameTBox.Text;
                defaultSetting.Default.department = departmentCbox.Text;
                defaultSetting.Default.target_path_d=savePathTextBox.Text;
                defaultSetting.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("設定儲存失敗 "+ex.Message);
            }
            closeWindow();
        }

        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            closeWindow();
        }

        private void dailyReportButton_Click(object sender, RoutedEventArgs e)
        {
            dailyReportTbox.Text = selectXlsFile();
        }

        private void weeklyReportButton_Click(object sender, RoutedEventArgs e)
        {
            weeklyReportTbox.Text= selectXlsFile();
        }

        private void workHourReportButton_Click(object sender, RoutedEventArgs e)
        {
            workHourReportTbox.Text=selectXlsFile();
        }
        /// <summary>
        /// 選取xls檔案
        /// </summary>
        /// <returns>檔案路徑</returns>
        private string selectXlsFile()
        {
            string filePath = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel File|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
            }
            return filePath;
        }
        /// <summary>
        /// 關閉本視窗
        /// </summary>
        private void closeWindow()
        {
            this.Close();
        }

        private void savePathButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFolderDialog openFolderDialog = new OpenFolderDialog()
            {
                Title = "請選擇目標位置",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                Multiselect = false
            };
            if (openFolderDialog.ShowDialog() == true)
            {
                savePathTextBox.Text = openFolderDialog.FolderName;
            }
        }
    }
}
