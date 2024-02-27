using JDI_ReportMaker.Util;
using NPOI.SS.UserModel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDI_ReportMaker.ExcelWriter
{
    class WorkHourExcelWritter:ExcelWritter
    {
        private IWorkbook workHourWorkBook;
        private SourceController sourceController;
        private readonly int[][] allSettingCell = { 
        new int[] {0,1,0 }, //幾年幾月工時表
        new int[] { 0,2,0}, //填表人:伍佑群
        new int[]{0,3,0},//統計期間:
        new int[]{3,1,0 },//填表人:伍佑群
        new int[] {3,2,32 }};//(直接寫入日期無前綴)統計期間
        public WorkHourExcelWritter(SourceController sourceController)
        {
            this.sourceController = sourceController;
            workHourWorkBook = sourceController.GetWorkBook(FileNameEnum.工時表);
        }
        public void WriteExcel()
        {
            WriteSetting();
            sourceController.saveFile(workHourWorkBook, FileNameEnum.工時表);
        }
        public override void WriteExcel(int[][] cells,IWorkbook target)
        {

        }
        private void WriteSetting()
        {
            string[] targetValue = GetCellTargetValueArr();
            for(int i=0;i<allSettingCell.Length;i++)
            {
                int[] targetCell = allSettingCell[i];
                overwriteFile(workHourWorkBook, targetCell[0], targetCell[1], targetCell[2], targetValue[i]);
            }
        }
        /// <summary>
        /// 獲得寫入設定之值陣列
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string[] GetCellTargetValueArr()
        {
            string[] valueArr=new string[allSettingCell.Length];
            string staffName = defaultSetting.Default.staff_name;
            string date = GetYYYMM();
            string dateDuration = GetThisMonthDuration();

            valueArr[0] = date+"工時表";
            valueArr[1] = "填表人: " + staffName;
            valueArr[2] = "統計期間: " + dateDuration;
            valueArr[3] = valueArr[1];
            valueArr[4] = dateDuration;
            return valueArr;
        }
        /// <summary>
        /// 獲得這個月期間之字串
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string GetThisMonthDuration()
        {
            string[] datePart = defaultSetting.Default.date.Split('-');
            string year = GetTaiwanYear(datePart[0]);
            string month = datePart[1];
            string lastDay = GetLastDayOfMonth(datePart[0],month);
            string result = $"{year}/{month}/01~{year}/{month}/{lastDay}";
            return result ;
        }
        /// <summary>
        /// 獲得當月最後一天
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string GetLastDayOfMonth(string yearStr, string monthStr)
        {
            int year=int.Parse(yearStr);
            int month=int.Parse(monthStr);  
            DateTime date=new DateTime(year,month,1);
            date=date.AddMonths(1);
            date=date.AddDays(-1);
            return date.ToString("dd");
        }

        /// <summary>
        /// 將緩存之日期轉為xxx年xx月
        /// </summary>
        /// <returns></returns>
        private string GetYYYMM()
        {
            string[] datePart = defaultSetting.Default.date.Split('-');
            string year = GetTaiwanYear(datePart[0]);
            string month = datePart[1];
            return  year+ "年" + month+"月";
        }
        /// <summary>
        /// 將西元年轉為民國年
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string GetTaiwanYear(string ceYearStr)
        {
            int ceYear=int.Parse(ceYearStr);
            int taiwanYear = ceYear - (2024 - 113);
            return taiwanYear.ToString();
        }
    }
}
