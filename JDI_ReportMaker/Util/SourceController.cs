using JDI_ReportMaker.ExcelWriter;
using JDI_ReportMaker.Util.ExcelWriter;
using JDI_ReportMaker.Util.PanelComponent;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace JDI_ReportMaker.Util
{
    /// <summary>
    /// excel操作、設定確認
    /// </summary>
    class SourceController
    {
        //用於命名檔案
        private readonly Dictionary<string, string> fileName
            = new Dictionary<string, string> { {"日報表", "佳帝科技員工日報表 "},
                                               {"週報表", "週報表 "},
                                               {"工時表", "工時表 "}};
        private Dictionary<string, string>? projectMap;
        private DBController? dbController;

        public SourceController(DBController dBController)
        {
            this.dbController = dBController;
            projectMap=GetJobProjectMap();
        }

        /// <summary>
        /// 確認來源檔案、設定是否正常
        /// </summary>
        /// <returns></returns>
        public bool SourceCheck()
        {
            try
            {
                return inputCheck() && FileCheck();
            }catch(Exception e) { MessageBox.Show(e.Message); return false; }
        }
        /// <summary>
        /// 確認員工姓名、編號、來源檔案、目的地是否有輸入
        /// </summary>
        /// <returns></returns>
        private bool inputCheck()
        {
            bool settingCheck = defaultSetting.Default.staff_name.Length > 0 &&
                defaultSetting.Default.department.Length > 0 &&
                defaultSetting.Default.source_path_d.Length > 0 &&
                defaultSetting.Default.source_path_w.Length > 0 &&
                defaultSetting.Default.source_path_h.Length > 0 &&
                defaultSetting.Default.target_path_d.Length > 0;
            if (settingCheck) { return true; }
            else
            {
                MessageBox.Show("請確認資料");
                return false;
            }
        }
        /// <summary>
        /// 用excel檔案中標題 確認來源檔案是否正常
        /// </summary>
        /// <returns></returns>
        private bool FileCheck()
        {
            IWorkbook target;
            if (!File.Exists(defaultSetting.Default.source_path_d))
            {
                MessageBox.Show("日報表檔案遺失");
                return false;
            }
            if (!File.Exists(defaultSetting.Default.source_path_w))
            {
                MessageBox.Show("週報表檔案遺失");
                return false;
            }
            if (!File.Exists(defaultSetting.Default.source_path_h))
            {
                MessageBox.Show("工時表檔案遺失");
                return false;
            }
            target = loadFile(defaultSetting.Default.source_path_d);
            if (target==null||target.GetSheetName(0) != "日報表")
            {
                MessageBox.Show(target==null?"請關閉所有報表":"日報表檔案異常");
                return false;
            }
            target = loadFile(defaultSetting.Default.source_path_w);
            if (target == null || target.GetSheetName(0) != "週報表")
            {
                MessageBox.Show(target == null ? "請關閉所有報表" : "日報表檔案異常");
                return false;
            }
            target = loadFile(defaultSetting.Default.source_path_h);
            if (target == null || target.GetSheetName(0) != "工時表")
            {
                MessageBox.Show(target == null ? "請關閉所有報表" : "日報表檔案異常");
                return false;
            }
            return true;
        }




        /// <summary>
        /// 寫入日報表
        /// </summary>
        /// <param name="panels"></param>
        /// <param name="tomorrowPanels"></param>
        public void WritePanelToExcel(List<TodayReportPanel> panels,List<TomorrowReportPanel> tomorrowPanels)
        {
            FileNameEnum targetType = FileNameEnum.日報表;
            try
            {
                IWorkbook target= staffDataWrite(targetType);
                if(target != null )
                {
                    //將面板內容寫入Excel
                    DailyReportWriter dailyReportWriter = new DailyReportWriter(this);
                    dailyReportWriter.WriteTodayPanel(panels, target);
                    dailyReportWriter.WriteTomorrowPanel(tomorrowPanels, target);
                    saveFile(target, targetType);
                }

            }
            catch { throw; }
        }

        /// <summary>
        /// 檢查面板內容並彈出提醒視窗
        /// </summary>
        /// <param name="panels"></param>
        /// <returns></returns>
        public bool CheckPanelInput(List<TodayReportPanel> panels)
        {
            //檢查日工時 報表內容
            int totalWorkHourToday = 0;
            string problemList="";
            foreach(TodayReportPanel panel in panels)
            {
                string index=panel.GetPanelNum();
                if (panel.GetProjectName() == "未選擇專案")
                {
                    problemList += $"第{index}欄尚未選擇專案項目\n";
                }
                if (panel.GetTitle().Length == 0)
                {
                    problemList += $"第{index}欄未輸入大項列表\n";
                }
                if (panel.GetDescribtion().Length == 0)
                {
                    problemList += $"第{index}欄未輸入細項說明\n";
                }
                if (panel.GetWorkHour().Equals("0"))
                {
                    problemList += $"第{index}欄未輸入工時\n";
                }
                totalWorkHourToday +=int.Parse(panel.GetWorkHour());
            }
            //工時總結異常時
            if(totalWorkHourToday !=8)
            {
                problemList += "工時加總不是8小時\n";
            }
            return problemList.Length == 0 ? true : WarningBox(problemList);

        }
        public bool CheckPanelInput(List<TomorrowReportPanel> panels)
        {
            string problemList = "";
            foreach (TomorrowReportPanel panel in panels)
            {
                string index = panel.GetPanelNum();

                if (panel.GetTitle().Length == 0)
                {
                    problemList += $"第{index}欄明日預計尚未輸入大項列表\n";
                }
                if (panel.GetDescribtion().Length == 0)
                {
                    problemList += $"第{index}欄明日預計尚未輸入細項\n";
                }
            }
            return problemList.Length == 0 ? true : WarningBox(problemList);

        }
        /// <summary>
        /// 可自定義文字的警告視窗
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        internal bool WarningBox(string message)
        {
            MessageBoxResult result=MessageBox.Show("問題列表:\n"+message+"\n 是否繼續執行?","警告",MessageBoxButton.YesNo,MessageBoxImage.Warning);
            if(result == MessageBoxResult.Yes)
            {
                return true;
            }
            return false;
        }
        /// <summary>
        /// 將員工姓名、日期寫入
        /// </summary>
        internal IWorkbook staffDataWrite(FileNameEnum fileNameEnum)
        {
            int[][] cellPath = new int[3][];
            ExcelWritter? settingWritter;
            IWorkbook? target;
            switch (fileNameEnum)
            {
                case FileNameEnum.日報表:
                    target = loadFile(defaultSetting.Default.source_path_d);
                    settingWritter = new DailyWeeklySettingWirter();
                    cellPath[0] = [0, 5, 0];//姓名
                    cellPath[1] = [0, 5, 3];//部門
                    cellPath[2] = [0, 5, 7]; break; //日期
                case FileNameEnum.週報表:
                    target = loadFile(defaultSetting.Default.source_path_w);
                    settingWritter = new DailyWeeklySettingWirter();
                    cellPath[0] = [0, 5, 0];//姓名
                    cellPath[1] = [0, 5, 4];//部門 
                    cellPath[2] = [0, 5, 7]; break;//日期
                case FileNameEnum.工時表:
                    target = loadFile(defaultSetting.Default.source_path_h);
                    settingWritter = new WorkHourExcelWritter();
                    cellPath[0] = [0, 1, 0];//yyyy年mm月工時表
                    cellPath[1] = [0, 2, 0];//填表人:姓名
                    cellPath[2] = [0, 3, 0];//統計期間
                    cellPath[3] = [3, 1, 0];//填表人:姓名
                    cellPath[4] = [3, 2, 32]; break;//(直接寫入日期無前綴)統計期間
                default: 
                    settingWritter = null;
                    target = null;
                    break;
            }
            if (settingWritter != null&& target!=null)
            {
                settingWritter.WriteExcel(cellPath, target);
            }
            return target;
        }
        /// <summary>
        /// 讀取目標報表
        /// </summary>
        /// <param name="sourcePath">目標報表之路徑</param>
        /// <returns>回傳讀取完之報表</returns>
        private IWorkbook loadFile(string sourcePath)
        {
            IWorkbook sourceWorkbook;
            try
            {
                using (FileStream file = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
                {
                    sourceWorkbook = new HSSFWorkbook(file);
                }
                return sourceWorkbook;
            }
            catch(Exception e)
            {
                MessageBox.Show("檔案路徑異常或目標檔案正開啟\n"+e.Message);
                return null;
            }
        }
        /// <summary>
        /// 另存xls檔
        /// </summary>
        /// <param name="sourceWorkbook">原檔</param>
        /// <param name="targerFilePath">另存檔案的位置與名稱</param>
        internal void saveFile(IWorkbook sourceWorkbook, FileNameEnum fileNameEnum)
        {
            //設定檔名:報表名+員工名+日期
            string? date = defaultSetting.Default.date;
            string fileNameStr = "/" + fileName[fileNameEnum.ToString()] +
                " " + defaultSetting.Default.staff_name +
                " " + date + ".xls";
            string targetFilePath = defaultSetting.Default.target_path_d + fileNameStr;
            try
            {
                using (FileStream file = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                {
                    sourceWorkbook.Write(file);
                }
            }
            catch
            {
                throw;
            }
        }
        /// <summary>
        /// 獲取工時表的專案清單<名稱,編號>
        /// </summary>
        /// <returns></returns>
        public Dictionary<string,string> GetJobProjectMap()
        {
            Dictionary<string,string> jobProjects = new Dictionary<string,string>();
            //讀取工時表的專案
            IWorkbook workHourWorkbook = loadFile(defaultSetting.Default.source_path_h);
            //找到報價單號分頁
            if (workHourWorkbook == null)
            {
                return jobProjects;
            }
            ISheet workHourSheet = workHourWorkbook.GetSheet("報價單號-專案名稱(未完成)");
            if (workHourSheet == null) workHourSheet = workHourWorkbook.GetSheetAt(1);
            for(int i = 2;i<workHourSheet.LastRowNum;i++)
            {
                IRow workHourRow=workHourSheet.GetRow(i);
                if (workHourRow.GetCell(1).ToString().Length > 0)
                {
                    jobProjects.Add(workHourRow.GetCell(1).ToString(), workHourRow.GetCell(0).ToString());
                }
                else { break; }
            }
            return jobProjects;
        }
        /// <summary>
        /// 獲取資料表
        /// </summary>
        /// <param name="fileType"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        internal ISheet GetReportSheet(FileNameEnum fileType,IWorkbook target)
        {
            ISheet sheet =target.GetSheet(fileType.ToString());
            if(sheet == null)
            {
                sheet= target.GetSheetAt(0);
            }
            return sheet;
        }

        internal DBController GetDBController()
        {
            return dbController;
        }
    }
}
