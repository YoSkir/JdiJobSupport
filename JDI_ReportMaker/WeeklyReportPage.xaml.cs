using JDI_ReportMaker.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
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
    /// WeeklyReportPage.xaml 的互動邏輯
    /// </summary>
    public partial class WeeklyReportPage : Window
    {
        private SourceController? sourceController;
        private DateTime? date;
        private DateTime[]? thisWeek;
        private int thisMonth = DateTime.Now.Month;
        private int thisYear = DateTime.Now.Year;

        public WeeklyReportPage()
        {
            InitializeComponent();
            sourceController = new SourceController();
            SetWeekComboBox();
            date = DateTime.Now;
            thisWeek =new DateTime[7];
        }

        private void SetWeekComboBox()
        {
            weekCbox.IsEditable = true;
            weekCbox.Text = "請選擇期間(默認本週)";
            weekCbox.ItemsSource=GetWeekList();
        }

        private List<string> GetWeekList()
        {
            List<string> weekList = new List<string>();

            int weekCount = 1;
            DateTime date = new DateTime(thisYear, thisMonth, 1);
            while (date.DayOfWeek != DayOfWeek.Monday)
            {
                date = date.AddDays(1);
            }
            while (date.Month == thisMonth)
            {
                string week = date.ToString("yyyy-MM-dd") + "~" + date.AddDays(6).ToString("MM-dd")+$" 第{weekCount}週";
                weekList.Add(week);
                date = date.AddDays(7);
                weekCount++;
            } 
            return weekList;
        }
        /// <summary>
        /// 寫入週報表的月曆
        /// </summary>
        private void DrawCalender(IWorkbook weekReport)
        {
            ISheet sheet = weekReport.GetSheet("週報表");
            if (sheet == null) sheet = weekReport.GetSheetAt(0);
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

        private void SetThisWeekArr(DateTime day)
        {
            while(day.DayOfWeek != DayOfWeek.Monday)
            {
                day = day.AddDays(-1);
            }
            for(int i = 0; i < 7; i++) 
            {
                thisWeek[i]= day;
                day = day.AddDays(1);
            }
        }

        private void saveWeekReportButton_Click(object sender, RoutedEventArgs e)
        {
            if(sourceController == null) { sourceController=new SourceController(); }
            IWorkbook target = sourceController.staffDataWrite(FileNameEnum.周報表);
            DrawCalender(target);
            sourceController.ExcecuteFile(target);
        }
    }
}
