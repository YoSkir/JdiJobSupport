using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDI_ReportMaker
{
    interface WriteExcel
    {
        void WriteExcel(int[][] cells, IWorkbook target);
        void overwriteFile(IWorkbook target, int sheet, int row, int cell, string value)
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
            catch
            {
                throw;
            }
        }
    }
}
