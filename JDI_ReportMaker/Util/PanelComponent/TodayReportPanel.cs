using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace JDI_ReportMaker.Util.PanelComponent
{
    internal class TodayReportPanel : StackPanelComponent
    {
        private StackPanel? thisPanel;
        private Label? numLabel;
        private ComboBox? comboBox;
        private TextBox? projectTitle, projectDescription, hourSpent;
        private CheckBox? doneCheck;
        private Button? addBtn, removeBtn;
        private MainWindow parentWindow;

        private Dictionary<string, string>? projectsNameCode;

        public TodayReportPanel(MainWindow mainWindow)
        {
            parentWindow = mainWindow;
            //設定專案清單
            SetJobProjectList();
            //創造面板
            CreatePanel();
            //連接主畫面
        }
        public override StackPanel GetPanel()
        {
            if (thisPanel != null)
            {
                return thisPanel;
            }
            return CreatePanel();
        }
        /// <summary>
        /// 創建一個於視窗上顯示輸入的面板欄位
        /// </summary>
        public override StackPanel CreatePanel()
        {
            if(thisPanel == null)
            {
                thisPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 10, 0, 0)
                };
            }
            //設定面板的原件
            numLabel = new Label();
            comboBox = new ComboBox { Margin = new Thickness(10, 0, 0, 0), Width = 281, Height = 32 };
            projectTitle = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 386, Height = 32, FontSize = 20 };
            projectDescription = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 221, Height = 32, FontSize = 20 };
            hourSpent = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 50, FontSize = 20 };
            doneCheck = new CheckBox { Margin = new Thickness(20, 0, 0, 0), Height = 32, Content = "已完成" };
            addBtn = new Button { Margin = new Thickness(50, 0, 0, 0), Content = "+" };
            removeBtn = new Button { Margin = new Thickness(10, 0, 0, 0), Content = "-" };
            //面板
            thisPanel.Children.Add(numLabel);
            //工作專案下拉選單
            if (projectsNameCode == null)
            {
                SetJobProjectList();
            }
            comboBox.ItemsSource = projectsNameCode.Keys;
            comboBox.IsEditable = true;
            comboBox.Text = "工時表統計用";
            thisPanel.Children.Add(comboBox);
            //大項列表、備註、工時、完成選格
            SetInputTip("請輸入大項列表", projectTitle);
            thisPanel.Children.Add(projectTitle);
            SetInputTip("請輸入備註", projectDescription);
            thisPanel.Children.Add(projectDescription);
            SetInputTip("工時", hourSpent);
            thisPanel.Children.Add(hourSpent);
            thisPanel.Children.Add(doneCheck);
            //面板增減按鈕
            addBtn.Click += AddButton_Clicked;
            thisPanel.Children.Add(addBtn);
            removeBtn.Click += RemoveButton_Clicked;
            thisPanel.Children.Add(removeBtn);
            return thisPanel;
        }
        private void SetJobProjectList()
        {
            try
            {
                SourceController sourceController = parentWindow.GetSourceController();
                projectsNameCode = sourceController.GetJobProjectMap();
            }
            catch (Exception ex) { MessageBox.Show("專案列表載入失敗，請確認工時表是否正常"); }
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
        /// <summary>
        /// comboBox選擇改變時，改變大項內容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Source_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //目前會延遲換
            projectTitle.Text = comboBox.Text;
            projectTitle.Background = new SolidColorBrush(Colors.Transparent);
        }
        private void AddButton_Clicked(object sender, RoutedEventArgs e)
        {
            parentWindow.AddTodayPanel();
        }
        private void RemoveButton_Clicked(object sender, RoutedEventArgs e)
        {
            parentWindow.RemovePanel(this);
        }

        public string GetPanelNum() { return numLabel.Content.ToString(); }
        public string GetTitle() { return projectTitle.Text; }
        public string GetProjectName() { return comboBox.Text; }
        public string GetProjectCode() { return projectsNameCode[comboBox.Text]; }
        public string GetDescribtion() { return projectDescription.Text; }
        public string GetWorkHour()
        {
            if(hourSpent.Text.Length == 0)
            {
                return "0";
            }
            return hourSpent.Text;
        }
        public bool GetDone() { return doneCheck.IsChecked == true; }
    }
}
