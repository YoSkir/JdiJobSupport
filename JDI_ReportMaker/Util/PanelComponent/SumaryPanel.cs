using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace JDI_ReportMaker.Util.PanelComponent
{
    internal class SumaryPanel : StackPanelComponent
    {
        private StackPanel thisPanel=new StackPanel();
        private Label projectCode=new Label();
        private Label projectName=new Label();
        private Label hourSpent=new Label();
        private CheckBox projectTeamOrManager=new CheckBox();

        private WorkHourReportPage? parantPage;
        private List<SumaryPanel>? sumaryPanelList;

        public SumaryPanel(WorkHourReportPage? parantPage)
        {
            CreatePanel();
            this.parantPage = parantPage;
            sumaryPanelList = parantPage.sumaryPanelList;
        }

        public override void AddPanel()
        {
            sumaryPanelList.Add(this);
        }

        public override StackPanel CreatePanel()
        {
            thisPanel.Margin = new Thickness(0, 10, 0, 0);
            thisPanel.Orientation = Orientation.Horizontal;
            projectCode.Width = 100;
            projectName.Width = 462;
            hourSpent.Width = 25;
            projectTeamOrManager.Content = Const.ProjectTeamManagerCheckBox;

            thisPanel.Children.Add(projectCode);
            thisPanel.Children.Add(projectName);
            thisPanel.Children.Add(hourSpent);
            thisPanel.Children.Add(projectTeamOrManager);

            return thisPanel;
        }

        public override StackPanel GetPanel()
        {
            if(thisPanel == null)return CreatePanel();
            return thisPanel;
        }

        public override void RemovePanel()
        {
            sumaryPanelList.Remove(this);
        }

        public void SetPanelValue(string projectCode,string projectName,string hourSpent)
        {
            this.projectCode.Content=projectCode;
            this.projectName.Content=projectName;
            this.hourSpent.Content=hourSpent;
        }
    }
}
