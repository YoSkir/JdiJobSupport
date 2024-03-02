using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace JDI_ReportMaker.ExcelWriter
{
    abstract class ExcelWritter
    {
        /// <summary>
        /// 將資料寫入指定Cells
        /// </summary>
        /// <param name="cells">[數量][sheet,row,cell]</param>
        /// <param name="target">目標檔案</param>
        public abstract void WriteExcel(int[][] cells, IWorkbook target);

        /// <summary>
        /// 檔案資料單格寫入(取代原有內容)
        /// </summary>
        /// <param name="target">目標檔案</param>
        /// <param name="sheet">工作表</param>
        /// <param name="row">列</param>
        /// <param name="cell">格</param>
        /// <param name="value">指定內容</param>
        public void overwriteFile(IWorkbook target, int sheet, int row, int cell, string value)
        {
            try
            {
                ISheet isheet = target.GetSheetAt(sheet);
                if (isheet != null)
                {
                    IRow irow = isheet.GetRow(row);
                    if (irow != null)
                    {
                        ICell icell = irow.GetCell(cell);
                        if (icell != null)
                        {
                            icell.SetCellValue(value);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        public void overwriteFile(IWorkbook target, int sheet, int row, int cell, int value)
        {
            try
            {
                ISheet isheet = target.GetSheetAt(sheet);
                if (isheet != null)
                {
                    IRow irow = isheet.GetRow(row);
                    if (irow != null)
                    {
                        ICell icell = irow.GetCell(cell);
                        if (icell != null)
                        {
                            icell.SetCellValue(value);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        /// <summary>
        /// 檔案資料單格寫入(不取代原有內容)
        /// </summary>
        /// <param name="target">目標檔案</param>
        /// <param name="sheet">工作表</param>
        /// <param name="row">列</param>
        /// <param name="cell">格</param>
        /// <param name="value">指定內容</param>
        private void writeFile(IWorkbook target, int sheet, int row, int cell, string value)
        {
            try
            {
                ISheet isheet = target.GetSheetAt(sheet);
                if (isheet != null)
                {
                    IRow irow = isheet.GetRow(row);
                    if (irow != null)
                    {
                        ICell icell = irow.GetCell(cell);
                        if (icell != null)
                        {
                            icell.SetCellValue(icell.StringCellValue + " " + value);
                        }
                    }
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
