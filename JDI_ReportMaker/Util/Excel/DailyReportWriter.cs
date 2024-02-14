using JDI_ReportMaker.ExcelWriter;
using JDI_ReportMaker.Util.PanelComponent;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;

namespace JDI_ReportMaker.Util.ExcelWriter
{
    class DailyReportWriter : ExcelWritter
    {
        SourceController controller = new SourceController();

        public override void WriteExcel(int[][] cells, IWorkbook target)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 將面板的資料寫入xls檔中
        /// </summary>
        /// <param name="panels"></param>
        /// <param name="target"></param>
        public void WriteTodayPanel(List<TodayReportPanel> panels,IWorkbook target)
        {
            ISheet sheet = controller.GetReportSheet(FileNameEnum.日報表, target);
            int rowIndex = 7;
            //設定勾選格的對齊方式、字體
            ICellStyle style = target.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            IFont font=target.CreateFont();
            font.Color = HSSFColor.Black.Index;
            font.FontHeightInPoints = 10;
            style.SetFont(font);
            try
            {
                for(int row=0;row< panels.Count; row++)
                {
                    IRow panelRow= sheet.GetRow(rowIndex+row);
                    panelRow.GetCell(1).SetCellValue(panels[row].GetPanelNum());
                    panelRow.GetCell(2).SetCellValue(panels[row].GetTitle());
                    //刪除原本的勾選框
                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.DrawingPatriarch;
                    var children = patriarch.Children;
                    var shapes = new List<HSSFShape>(children);
                    for(int i=0;i<shapes.Count;i++)
                    {
                        patriarch.RemoveShape(shapes[i]);
                    }
                    panelRow.GetCell(3).CellStyle = style;
                    if (panels[row].GetDone())
                    {
                        panelRow.GetCell(3).SetCellValue("\u2611 已完成            \u2610 未完成");
                    }
                    else
                    {
                        panelRow.GetCell(3).SetCellValue("\u2610 已完成            \u2611 未完成");

                    }
                    panelRow.GetCell(7).SetCellValue(panels[row].GetDescribtion());
                }
            }catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void WriteTomorrowPanel(List<TomorrowReportPanel> panels,IWorkbook target)
        {
            ISheet sheet = controller.GetReportSheet(FileNameEnum.日報表, target);
            int rowIndex = 15;
            for(int i=0;i< panels.Count;i++)
            {
                IRow row=sheet.GetRow(rowIndex+i);
                row.GetCell(1).SetCellValue(panels[i].GetPanelNum());
                row.GetCell(2).SetCellValue(panels[i].GetTitle());
                row.GetCell(3).SetCellValue(panels[i].GetDescribtion());
            }
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
