using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace JDI_ReportMaker.Util.PanelComponent
{
    public class TomorrowReportPanel:StackPanelComponent
    {
        private StackPanel? thisPanel;
        private Label? numLabel;
        private TextBox? projectTitle, projectDescription;
        private Button? addBtn, removeBtn;
        private MainWindow parentWindow;

        private List<TomorrowReportPanel> tomorrowPanels;

        public TomorrowReportPanel(MainWindow mainWindow)
        {
            //創造面板
            CreatePanel();
            //連接主畫面
            parentWindow = mainWindow;
            tomorrowPanels=parentWindow.GetTomorrowPanels();
        }

        public void SetPanelList(List<TomorrowReportPanel> panelList)
        {
            tomorrowPanels = panelList;
        }
        public override StackPanel CreatePanel()
        {
            if (thisPanel == null)
            {
                thisPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 10, 0, 0)
                };
            }
            //設定面板元件
            numLabel = new Label();
            projectTitle = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 386, Height = 32, FontSize = 20 };
            projectDescription = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 386, Height = 32, FontSize = 20 };
            addBtn = new Button { Margin = new Thickness(50, 0, 0, 0), Content = "+" };
            removeBtn = new Button { Margin = new Thickness(10, 0, 0, 0), Content = "-" };
            //面板
            thisPanel.Children.Add(numLabel);
            //輸入項
            SetInputTip("請輸入大項列表", projectTitle);
            thisPanel.Children.Add(projectTitle);
            SetInputTip("請輸入預計情況", projectDescription);
            thisPanel.Children.Add(projectDescription);
            //增減按鈕
            addBtn.Click += AddButton_Clicked;
            thisPanel.Children.Add(addBtn);
            removeBtn.Click += RemoveButton_Clicked;
            thisPanel.Children.Add(removeBtn);
            return thisPanel;
        }

        public override StackPanel GetPanel()
        {
            if(thisPanel != null)
            {
                return thisPanel;
            }
            else { return CreatePanel(); }
        }

        public override void AddPanel()
        {
            if (tomorrowPanels != null && tomorrowPanels.Count < 5)
            {
                TomorrowReportPanel panel = new TomorrowReportPanel(parentWindow);
                tomorrowPanels.Add(panel);
                parentWindow.ShowPanel();
            }
        }
        public override void RemovePanel()
        {
            if (tomorrowPanels != null && tomorrowPanels.Count > 1)
            {
                tomorrowPanels.Remove(this);
                parentWindow.ShowPanel();
            }
        }
        private void AddButton_Clicked(object sender, RoutedEventArgs e)
        {
            AddPanel();
        }
        private void RemoveButton_Clicked(object sender, RoutedEventArgs e)
        {
            RemovePanel();
        }
        /// <summary>
        /// 設定這個欄位的編號，由主視窗顯示時一併設定
        /// </summary>
        /// <param name="num"></param>
        public void SetPanelNum(int num)
        {
            if (numLabel != null)
                numLabel.Content = num.ToString();
        }
        public string GetPanelNum() { return numLabel.Content.ToString(); }
        public string GetTitle() { return projectTitle.Text; }

        public void SetTitle(string text)
        {
            projectTitle.Text= text;
        }
        public string GetDescribtion() { return projectDescription.Text; }
        public void SetDescribtion(string text)
        {
            projectDescription.Text= text;
        }
    }
}
