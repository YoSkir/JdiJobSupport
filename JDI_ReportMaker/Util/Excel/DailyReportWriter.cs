using JDI_ReportMaker.ExcelWriter;
using JDI_ReportMaker.Util.PanelComponent;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace JDI_ReportMaker.Util.ExcelWriter
{
    class DailyReportWriter : ExcelWritter
    {
        public override void WriteExcel(int[][] cells, IWorkbook target)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 將面板的資料寫入xls檔中
        /// </summary>
        /// <param name="panels"></param>
        /// <param name="target"></param>
        public void WritePanel(List<TodayReportPanel> panels,IWorkbook target)
        {
            ISheet sheet = target.GetSheet("日報表");
            if(sheet == null)
            {
                sheet = target.GetSheetAt(0);
            }
            int rowIndex = 7;
            try
            {
                for(int row=0;row< panels.Count; row++)
                {
                    IRow panelRow= sheet.GetRow(rowIndex+row);
                    panelRow.GetCell(1).SetCellValue(panels[row].GetPanelNum());
                    panelRow.GetCell(2).SetCellValue(panels[row].GetTitle());
                    if (panels[row].GetDone())
                    {
                        panelRow.GetCell(3).SetCellValue("\u00A3 未完成");
                    }
                    panelRow.GetCell(7).SetCellValue(panels[row].GetDescribtion());
                }
            }catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void ExcelCheckBox(bool jobDone,ISheet target,int panelIndex)
        {
            int targetCheckBox = (panelIndex - 1) * 2 + (jobDone? 0 : 1);
            //獲取cell的框選對象
            HSSFPatriarch patriarch = (HSSFPatriarch)target.DrawingPatriarch;
            HSSFShape shape = patriarch.Children[targetCheckBox];
        }
    }
}
