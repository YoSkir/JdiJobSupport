using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace JDI_ReportMaker.ExcelWriter
{
    class DailyWeeklySettingWirter : ExcelWritter
    {
        public override void WriteExcel(int[][] cells,IWorkbook target)
        {
            overwriteFile(target, cells[0][0], cells[0][1], cells[0][2], "姓名: " + defaultSetting.Default.staff_name);
            overwriteFile(target, cells[1][0], cells[1][1], cells[1][2], "部門: " + defaultSetting.Default.department);
            overwriteFile(target, cells[2][0], cells[2][1], cells[2][2], "日期: " + defaultSetting.Default.date);
        }
    }
}
