using JDI_ReportMaker.Util.PanelComponent;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDI_ReportMaker.Util
{
    public class WorkHourEntity
    {
        public string? recordDate {  get; set; }
        public string? projectCode { get; set; }
        public string? projectName { get; set;}
        public int hourSpent { get; set; }
        public int timePersent { get; set; }
        public bool managerOrTeam { get; set; }

    }
}
