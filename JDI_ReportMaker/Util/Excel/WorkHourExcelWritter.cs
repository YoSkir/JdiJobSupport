using JDI_ReportMaker.Util;
using NPOI.SS.UserModel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace JDI_ReportMaker.ExcelWriter
{
    class WorkHourExcelWritter:ExcelWritter
    {
        private IWorkbook workHourWorkBook;
        private SourceController sourceController;
        private MainWindow mainWindow;
        private readonly int[][] SETTING_CELL = { 
        new int[] {0,1,0 }, //幾年幾月工時表
        new int[] { 0,2,0}, //填表人:伍佑群
        new int[]{0,3,0},//統計期間:
        new int[]{3,1,0 },//填表人:伍佑群
        new int[] {3,2,32 }};//(直接寫入日期無前綴)統計期間
        private readonly int SUMARY_SHEET_INDEX = 0;
        private readonly int SUMARY_START_ROW = 6;
        private readonly int[] SUMARY_CELL = { 0, 1, 2 ,4};
        private readonly string PROJECT_IN_CHARGE_SYMBOL = "X";
        private readonly int[] TOTAL_WORKHOUR_PERSENT = { 0, 20, 2, 100 };
        private readonly int RECORD_START_ROW = 6;
        private readonly int RECORD_START_CELL = 2;
        private readonly int RECORD_SHEET_INDEX = 3;
        private readonly int RECORD_DATE_ROW = 4;
        private readonly int[] RECORD_PROJECT_CODE_NAME_CELL = { 0, 1 };


        public WorkHourExcelWritter(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
            this.sourceController = mainWindow.GetSourceController();
            workHourWorkBook = sourceController.GetWorkBook(FileNameEnum.工時表);
        }
        public void WriteExcel(string yearMonth)
        {
            WriteSetting(yearMonth);
            WriteSumary(yearMonth);
            WriteRecord(yearMonth);
            SaveWorkHourFile(yearMonth);
        }

        private void SaveWorkHourFile(string yearMonth)
        {
            string fileNameDate = GetFileNameDate(yearMonth);
            workHourWorkBook.GetSheetAt(RECORD_SHEET_INDEX).ForceFormulaRecalculation = true;
            sourceController.saveFile(workHourWorkBook, FileNameEnum.工時表, fileNameDate);
        }

        private string GetFileNameDate(string yearMonth)
        {
            string[] dateSplit = yearMonth.Split('-');
            return dateSplit[1] + "月";
        }

        private void WriteRecord(string yearMonth)
        {        
            Dictionary<string,int> projectName_rowIndex=new Dictionary<string,int>();
            Dictionary<string, int> date_cellIndex = new Dictionary<string, int>();
            List<WorkHourEntity> recordList = mainWindow.GetDBController().GetProjectRecordList(yearMonth);
            foreach(WorkHourEntity record in recordList)
            {
                int rowIndex = GetRowIndex(projectName_rowIndex, record);
                int cellIndex = GetCellIndex(date_cellIndex,record.recordDate);
                WriteOneRecord(rowIndex, cellIndex, record.hourSpent);
            }
        }

        private void WriteOneRecord(int rowIndex, int currentCellIndex, int hourSpent)
        {
            overwriteFile(workHourWorkBook, RECORD_SHEET_INDEX, rowIndex, currentCellIndex, hourSpent);
        }

        private int GetCellIndex(Dictionary<string,int> date_cellIndex,string recordDate)
        {
            recordDate = GetRecordFormatDate(recordDate);
            if (date_cellIndex.ContainsKey(recordDate))
            {
                return date_cellIndex[recordDate];
            }
            else
            {
                int cellIndex = RECORD_START_CELL + date_cellIndex.Count * 2;
                date_cellIndex.Add(recordDate, cellIndex);
                WriteRecordDate(cellIndex, recordDate);
                return cellIndex;
            }
        }

        private void WriteRecordDate(int currentCellIndex, string currentDate)
        {
            overwriteFile(workHourWorkBook,RECORD_SHEET_INDEX,RECORD_DATE_ROW,currentCellIndex,currentDate);
        }

        private string GetRecordFormatDate(string? recordDate)
        {
            string[] dateSplit = recordDate.Split(['/',' ']);
            string taiwanYear = GetTaiwanYear(dateSplit[0]);
            string result = taiwanYear + "/" + dateSplit[1] + "/" + dateSplit[2];
            return result;
        }

        private int GetRowIndex(Dictionary<string, int> projectName_rowIndex, WorkHourEntity record)
        {
            string projectName = record.projectName;
            if (projectName_rowIndex.ContainsKey(projectName))
            {
                return projectName_rowIndex[projectName];
            }
            else
            {
                int rowIndex = RECORD_START_ROW + projectName_rowIndex.Count;
                projectName_rowIndex.Add(projectName, rowIndex);
                WriteProjectCodeName(rowIndex, record);
                return rowIndex;
            }
        }
        /// <summary>
        /// 將專案編號、名稱寫入紀錄表
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="record"></param>
        private void WriteProjectCodeName(int rowIndex, WorkHourEntity record)
        {
            string[] projectInfo = { record.projectCode, record.projectName };
            for(int i = 0; i < 2; i++)
            {
                overwriteFile(workHourWorkBook, RECORD_SHEET_INDEX, rowIndex, RECORD_PROJECT_CODE_NAME_CELL[i], projectInfo[i]);
            }
        }


        /// <summary>
        /// 將報表統計寫入工時表
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        private void WriteSumary(string yearMonth)
        {
            List<WorkHourEntity> workHourList = mainWindow.GetWorkHourReportPage().GetWorkHourEntity(yearMonth);
            for(int i=0;i<workHourList.Count;i++)
            {
                WriteOneSumary(workHourList[i], i);
            }
            WriteTotalPersent();
        }
        /// <summary>
        /// 寫入總計時間比例
        /// </summary>
        private void WriteTotalPersent()
        {
            int sheet = TOTAL_WORKHOUR_PERSENT[0];
            int row = TOTAL_WORKHOUR_PERSENT[1];
            int cell = TOTAL_WORKHOUR_PERSENT[2];
            string value = TOTAL_WORKHOUR_PERSENT[3].ToString() + "%";
            overwriteFile(workHourWorkBook, sheet, row, cell, value);
        }

        /// <summary>
        /// 寫入一個統計面板
        /// </summary>
        /// <param name="workHourEntity"></param>
        /// <param name="currentIndex"></param>
        private void WriteOneSumary(WorkHourEntity workHourEntity, int currentIndex)
        {
            bool projectInCharge = mainWindow.GetWorkHourReportPage().GetProjectInCharge(workHourEntity.projectName);
            int currentRow = SUMARY_START_ROW + currentIndex;
            string[] dataToWrite = GetWorkHourDataStrArr(workHourEntity);
            for (int j = 0; j < dataToWrite.Length; j++)
            {
                //如果有勾選專案負責人，則cell index維持4，否則+1
                int projectInChargeAddOneIndex = 0;
                if (j == dataToWrite.Length - 1 && !projectInCharge)
                {
                    projectInChargeAddOneIndex = 1;
                }
                overwriteFile(workHourWorkBook, SUMARY_SHEET_INDEX, currentRow, SUMARY_CELL[j] + projectInChargeAddOneIndex, dataToWrite[j]);
            }
        }

        /// <summary>
        /// 從workhour容器獲得一次填入表單的字串陣列
        /// </summary>
        /// <param name="workHourEntity"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string[] GetWorkHourDataStrArr(WorkHourEntity workHourEntity)
        {
            string[] dataToWrite = { workHourEntity.projectCode, workHourEntity.projectName, 
                workHourEntity.timePersent.ToString()+"%",PROJECT_IN_CHARGE_SYMBOL};
            return dataToWrite;
        }

        public override void WriteExcel(int[][] cells,IWorkbook target)
        {

        }
        /// <summary>
        /// 將資料、日期寫入工時表
        /// </summary>
        private void WriteSetting(string yearMonth)
        {
            string[] targetValue = GetCellTargetValueArr(yearMonth);
            for(int i=0;i<SETTING_CELL.Length;i++)
            {
                int[] targetCell = SETTING_CELL[i];
                overwriteFile(workHourWorkBook, targetCell[0], targetCell[1], targetCell[2], targetValue[i]);
            }
        }
        /// <summary>
        /// 獲得寫入設定之值陣列
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string[] GetCellTargetValueArr(string yearMonth)
        {
            string[] valueArr=new string[SETTING_CELL.Length];
            string staffName = defaultSetting.Default.staff_name;
            string date = GetYYYMM(yearMonth);
            string dateDuration = GetThisMonthDuration(yearMonth);

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
        private string GetThisMonthDuration(string yearMonth)
        {
            string[] datePart = yearMonth.Split('-');
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
        private string GetYYYMM(string yearMonth)
        {
            string[] datePart = yearMonth.Split('-');
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
