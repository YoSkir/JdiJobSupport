using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace JDI_ReportMaker.Util.PanelComponent
{
    internal class SumaryPanel : StackPanelComponent
    {
        private StackPanel thisPanel=new StackPanel();
        private Label projectCode=new Label();
        private Label projectName=new Label();
        private Label hourSpent=new Label();
        private Label timePersent=new Label();
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
            //設計第一層
            StackPanel borderPanel = new StackPanel();
            Border outBorder=new Border();
            outBorder.BorderBrush = System.Windows.Media.Brushes.Black;
            outBorder.BorderThickness = new Thickness(1);
            outBorder.Height = 50;
            outBorder.Width = 486;
            borderPanel.Orientation = Orientation.Horizontal;
            outBorder.Child = borderPanel;
            //開始第二~四層設計
            secondLayerPanel(borderPanel);
            thisPanel.Children.Add(outBorder);
            return thisPanel;
        }
        /// <summary>
        /// 設定面板的值
        /// </summary>
        /// <param name="panelEntity"></param>
        public void SetPanelValue(WorkHourEntity panelEntity)
        {
            projectName.Content = panelEntity.projectName;
            projectCode.Content=panelEntity.projectCode;
            hourSpent.Content=panelEntity.hourSpent;
            timePersent.Content=panelEntity.timePersent+"%";
        }
        private void secondLayerPanel(StackPanel parentPanel)
        {
            StackPanel leftPanel=new StackPanel();leftPanel.Height= parentPanel.Height;leftPanel.Width = 352;
            //第三層
            thirdUpPanel(leftPanel);
            thirdDownPanel(leftPanel);
            StackPanel rightPanel=new StackPanel();rightPanel.Height= parentPanel.Height;rightPanel.Width= 134;rightPanel.VerticalAlignment 
                = VerticalAlignment.Center;
            CheckBox check = projectTeamOrManager; check.Content = Const.ProjectTeamManagerCheckBox;
            rightPanel.Children.Add(check);

            parentPanel.Children.Add(leftPanel);
            parentPanel.Children.Add(rightPanel);
        }

        private void thirdDownPanel(StackPanel secondLeftPanel)
        {
            StackPanel thirdDownPanel=new StackPanel();thirdDownPanel.Height = 25; thirdDownPanel.Orientation=Orientation.Horizontal;
            Label downLabelTitle = new Label();downLabelTitle.Content = "專案名稱";
            Label downLableValue = projectName;
            thirdDownPanel.Children.Add(downLabelTitle);
            thirdDownPanel.Children.Add(downLableValue);
            secondLeftPanel.Children.Add(thirdDownPanel);
        }

        private void thirdUpPanel(StackPanel secondLeftPanel)
        {
            StackPanel thirdUpPanel=new StackPanel(); thirdUpPanel.Height = 25; thirdUpPanel.Orientation = Orientation.Horizontal;
                StackPanel panelLeft=new StackPanel();panelLeft.Width = 150;panelLeft.Orientation = Orientation.Horizontal;
                    Label leftLabelTitle=new Label();leftLabelTitle.Content = "專案編號";
                    Label leftLabelValue = projectCode;
                StackPanel panelMid = new StackPanel();panelMid.Width = 100;panelMid.Orientation = Orientation.Horizontal;
                    Label midLabelTitle=new Label();midLabelTitle.Content = "總工時";
                    Label midLabelValue = hourSpent;
                StackPanel panelRight=new StackPanel();panelRight.Width = 100;panelRight.Orientation = Orientation.Horizontal;
                    Label rightLabelTitle=new Label();rightLabelTitle.Content = "工時比例:";
                    Label rightLabelValue = timePersent;
            panelLeft.Children.Add(leftLabelTitle);
            panelLeft.Children.Add(leftLabelValue);
            panelMid.Children.Add(midLabelTitle);
            panelMid.Children.Add(midLabelValue);
            panelRight.Children.Add(rightLabelTitle);
            panelRight.Children.Add(rightLabelValue);
            thirdUpPanel.Children.Add(panelLeft);
            thirdUpPanel.Children.Add(panelMid);
            thirdUpPanel.Children.Add(panelRight);
            secondLeftPanel.Children.Add(thirdUpPanel);
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

        public string GetProjectName()
        {
            return projectName.Content.ToString();
        }

        public bool GetProjectInCharge()
        {
            return projectTeamOrManager.IsChecked==true;
        }
    }
}
