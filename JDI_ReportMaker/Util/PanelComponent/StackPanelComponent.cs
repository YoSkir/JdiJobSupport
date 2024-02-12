using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace JDI_ReportMaker.Util.PanelComponent
{
    public abstract class StackPanelComponent
    {
        public abstract StackPanel GetPanel();
        public abstract void CreatePanel();
    }
}
