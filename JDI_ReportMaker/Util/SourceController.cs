using JDI_ReportMaker.ExcelWriter;
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
    class SourceController
    {
        private readonly Dictionary<string, string> fileName
            = new Dictionary<string, string> { {"日報表", "佳帝科技員工日報表 "},
                                               {"周報表", "佳帝科技員工周報表 "},
                                               {"工時表", "佳帝科技員工工時表 "}};
        /// <summary>
        /// 確認來源檔案、設定是否正常
        /// </summary>
        /// <returns></returns>
        public bool SourceCheck()
        {
            return inputCheck() && FileCheck();
        }
        public void ExcecuteFile()
        {
            FileNameEnum targetType = FileNameEnum.日報表;
            try
            {
                IWorkbook target= staffDataWrite(targetType);
                if(target != null )
                {
                    saveFile(target, targetType);
                }

            }
            catch { throw; }
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
            if(target.GetSheetName(0) != "日報表")
            {
                MessageBox.Show("日報表檔案異常");
                return false;
            }
            target = loadFile(defaultSetting.Default.source_path_w);
            if (target.GetSheetName(0) != "週報表")
            {
                MessageBox.Show("週報表檔案異常");
                return false;
            }
            target = loadFile(defaultSetting.Default.source_path_h);
            if (target.GetSheetName(0) != "工時表")
            {
                MessageBox.Show("工時表檔案異常");
                return false;
            }
            return true;
        }
        /// <summary>
        /// 確認員工姓名、編號、來源檔案、目的地是否有輸入
        /// </summary>
        /// <returns></returns>
        private bool inputCheck()
        {
            return defaultSetting.Default.staff_name.Length > 0 &&
                defaultSetting.Default.department.Length > 0 &&
                defaultSetting.Default.source_path_d.Length > 0 &&
                defaultSetting.Default.source_path_w.Length > 0 &&
                defaultSetting.Default.source_path_h.Length > 0 &&
                defaultSetting.Default.date.Length > 0 &&
                defaultSetting.Default.target_path_d.Length > 0;
        }
        /// <summary>
        /// 將員工姓名、日期寫入
        /// </summary>
        private IWorkbook staffDataWrite(FileNameEnum fileNameEnum)
        {
            int[][] cellPath = new int[3][];
            ExcelWritter? settingWritter;
            IWorkbook? target;
            switch (fileNameEnum)
            {
                case FileNameEnum.日報表:
                    target = loadFile(defaultSetting.Default.source_path_d);
                    settingWritter = new DailyWeeklyExcelWirter();
                    cellPath[0] = [0, 5, 0];//姓名
                    cellPath[1] = [0, 5, 3];//部門
                    cellPath[2] = [0, 5, 7]; break; //日期
                case FileNameEnum.周報表:
                    target = loadFile(defaultSetting.Default.source_path_w);
                    settingWritter = new DailyWeeklyExcelWirter();
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
            }
            catch
            {
                throw;
            }
            return sourceWorkbook;
        }
        /// <summary>
        /// 另存xls檔
        /// </summary>
        /// <param name="sourceWorkbook">原檔</param>
        /// <param name="targerFilePath">另存檔案的位置與名稱</param>
        private void saveFile(IWorkbook sourceWorkbook, FileNameEnum fileNameEnum)
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

    }
}
