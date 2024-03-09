using JDI_ReportMaker.Util;
using MahApps.Metro.Controls;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.CodeDom;
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
    /// WeeklyReportPage.xaml 的互動邏輯
    /// </summary>
    public partial class WeeklyReportPage : Window
    {
        private SourceController sourceController;
        //存放指定週每一天的陣列
        private DateTime[]? thisWeekEveryDay;
        //存放本月每週第一天的陣列
        private List<DateTime>? everyWeekMonday;
        private MainWindow? parent;
        private int thisMonth = DateTime.Now.Month;
        private int thisYear = DateTime.Now.Year;

        public WeeklyReportPage(MainWindow mainWindow)
        {
            InitializeComponent();
            SaveButtonLogLabel.Content = "";
            parent = mainWindow;
            sourceController = parent.GetSourceController();
            everyWeekMonday = new List<DateTime>();
            SetWeekComboBox();
            thisWeekEveryDay =new DateTime[7];
        }
        /// <summary>
        /// 初始化下拉選單與設定本月每週區間
        /// </summary>
        private void SetWeekComboBox()
        {
            weekCbox.IsEditable = true;
            weekCbox.Text = "請選擇期間(默認本週)";
            weekCbox.ItemsSource=GetWeekList(thisYear,thisMonth);
        }
        /// <summary>
        /// 獲得下拉選單要用的本月每個週間
        /// </summary>
        /// <returns></returns>
        private List<string> GetWeekList(int thisYear, int thisMonth)
        {
            List<string> weekList = new List<string>();

            int weekCount = 1;
            DateTime date = new DateTime(thisYear, thisMonth, 1);
            //因為每個月最後一週有可能到下個月的前幾天，所以從本月第一個禮拜一開始
            while (date.DayOfWeek != DayOfWeek.Monday)
            {
                date = date.AddDays(1);
            }
            //獲得本月每週區間與每週第一天
            if (everyWeekMonday == null) { everyWeekMonday = new List<DateTime>(); }
            while (date.Month == thisMonth)
            {
                //每週第一天
                everyWeekMonday.Add(date);
                //每週區間
                string week = date.ToString("yyyy-MM-dd") + "~" + date.AddDays(6).ToString("MM-dd")+$" 第{weekCount}週";
                //判斷本週是哪個週間 將本週週間文字加上"本週"
                if(DateTime.Today>=date&&DateTime.Today<date.AddDays(7)) { week += "(本週)"; }
                date = date.AddDays(7);
                weekList.Add(week);
                weekCount++;
            } 
            return weekList;
        }
        /// <summary>
        /// 寫入週報表的月曆
        /// </summary>
        private void DrawCalender(IWorkbook weekReport)
        {
            ISheet sheet = sourceController.GetReportSheet(FileNameEnum.週報表, weekReport);
            //設定月曆旁的本月份
            string callenderTitle = DateTime.Now.ToString("yyyy/MM");
            IRow row=sheet.GetRow(7);
            ICell cell = row.GetCell(0);
            cell.SetCellValue(callenderTitle+"月份");
            //畫月曆
            int currentRow = 8;
            DateTime date = new DateTime(thisYear,thisMonth,1);
            //清除1號以前的格子
            int currentCell = GetCallenderCellIndex(date.DayOfWeek)-1;
            row = sheet.GetRow(currentRow);
            while (currentCell > 1) { row.GetCell(currentCell).SetCellValue("");currentCell--; }
            //開始畫月曆
            currentCell = GetCallenderCellIndex(date.DayOfWeek);
            while (date.Month == thisMonth)
            {
                row = sheet.GetRow(currentRow);
                row.GetCell(currentCell).SetCellValue(date.Day);
                currentCell++;
                date=date.AddDays(1);
                if (currentCell > 8)
                {
                    currentCell = 2;
                    currentRow += 1;
                }
            }
            //清除月曆後部
            while (currentRow < 14)
            {
                row = sheet.GetRow(currentRow);
                row.GetCell(currentCell).SetCellValue("");
                currentCell++;
                if (currentCell > 8)
                {
                    currentCell = 2;
                    currentRow += 1;
                }
            }
        }
        /// <summary>
        /// 獲得週報表月曆星期對應的行
        /// </summary>
        /// <param name="dayOfWeek"></param>
        /// <returns></returns>
        private int GetCallenderCellIndex(DayOfWeek dayOfWeek)
        {
            switch(dayOfWeek)
            {
                case DayOfWeek.Sunday:return 8;
                case DayOfWeek.Monday:return 2;
                case DayOfWeek.Tuesday: return 3;
                case DayOfWeek.Wednesday: return 4;
                case DayOfWeek.Thursday: return 5;
                case DayOfWeek.Friday: return 6;
                case DayOfWeek.Saturday: return 7;
                default:return 9;
            }
        }
        /// <summary>
        /// 將本週每一天存入陣列
        /// </summary>
        /// <param name="day"></param>
        private void SetThisWeekArr()
        {
            DateTime day= DateTime.Today;
            //判斷如果選單有選擇特定週間
            if (weekCbox.SelectedIndex > -1)
            {
                day = everyWeekMonday[weekCbox.SelectedIndex];
            }
            while(day.DayOfWeek != DayOfWeek.Monday)
            {
                day = day.AddDays(-1);
            }
            for(int i = 0; i < 7; i++) 
            {
                if (thisWeekEveryDay == null) thisWeekEveryDay = new DateTime[7];
                thisWeekEveryDay[i]= day;
                day = day.AddDays(1);
            }
        }

        private void saveWeekReportButton_Click(object sender, RoutedEventArgs e)
        {
            FileNameEnum fileType = FileNameEnum.週報表;
            try
            {
                if (sourceController.SourceCheck())
                {
                    SetThisWeekArr();
                    IWorkbook target = sourceController.staffDataWrite(fileType);
                    DrawCalender(target);
                    WriteWeekReportDB();
                    WritePanelToExcel(target);
                    SaveWeeklyReport(target);
                    MessageBox.Show("儲存成功");
                    this.Close();
                }
            }
            catch { MessageBox.Show("儲存失敗"); }

        }

        private void WriteWeekReportDB()
        {
            throw new NotImplementedException();
        }

        private void SaveWeeklyReport(IWorkbook target)
        {
            string[] weekName = weekCbox.Text.Split(' ');
            string fileNameDate = weekName[weekName.Length - 1];
            sourceController.saveFile(target, FileNameEnum.週報表,fileNameDate);
        }

        /// <summary>
        /// 將面板內容寫入報表
        /// </summary>
        /// <param name="target"></param>
        private void WritePanelToExcel(IWorkbook target)
        {
            try
            {
                ISheet sheet = sourceController.GetReportSheet(FileNameEnum.週報表, target);
                IRow irow;
                //獲得面板上內容的陣列(由左至右由上至下共12格)
                string[] text = GetPanelText();
                for (int row=0;row<6; row++)
                {
                    irow = sheet.GetRow(17 + row*2);
                    //日期格
                    irow.GetCell(0).SetCellValue(thisWeekEveryDay[row].ToString("MM/dd"));
                    //工作計畫
                    irow.GetCell(2).SetCellValue(text[4*row]);
                    //內容說明
                    irow.GetCell(6).SetCellValue(text[4 * row+1]);
                    //計畫與說明的第二行
                    if (row < 5)//排除週六
                    {
                        irow = sheet.GetRow(17 + row * 2 + 1);
                        //工作計畫
                        irow.GetCell(2).SetCellValue(text[4 * row+2]);
                        //內容說明
                        irow.GetCell(6).SetCellValue(text[4 * row + 3]);
                    }
                }
            }catch (Exception ex) { MessageBox.Show("週報表面板寫入失敗\n"+ex.Message); }
        }
        /// <summary>
        /// 獲得面板上內容的陣列(由左至右由上至下共12格)
        /// </summary>
        /// <returns></returns>
        private string[] GetPanelText()
        {
            string[] text = new string[22];
            int arrIndex = 0;
            foreach(UIElement element in weekRoportGrid.Children)
            {
                if(element is TextBox)
                {
                    text[arrIndex] = (element as TextBox).Text;
                    arrIndex++;
                }
            }
            return text;
        }
    }
}
