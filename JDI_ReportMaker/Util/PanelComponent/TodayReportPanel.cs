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
        private TextBox? projectTitle, projectDescription,hourSpent;
        private CheckBox? doneCheck;
        private Button? addBtn, removeBtn;
        private MainWindow parentWindow;

        public TodayReportPanel(MainWindow mainWindow)
        {
            CreatePanel();
            parentWindow=mainWindow;
        }
        public override StackPanel GetPanel()
        {
            if(thisPanel != null)
            {
                return thisPanel;
            }
            return new StackPanel();
        }
        /// <summary>
        /// 創建一個於視窗上顯示輸入的面板欄位
        /// </summary>
        public override void CreatePanel()
        {
            thisPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 10, 0, 0)
            };
            numLabel = new Label();
            comboBox = new ComboBox { Margin = new Thickness(10, 0, 0, 0), Width=281,Height=32};
            projectTitle = new TextBox { Margin = new Thickness(10, 0, 0, 0) ,Width=386,Height=32,FontSize=20};
            projectDescription = new TextBox { Margin = new Thickness(10, 0, 0, 0) ,Width=221,Height=32, FontSize = 20 };
            hourSpent = new TextBox { Margin = new Thickness(10, 0, 0, 0), Width = 50 , FontSize = 20 };
            doneCheck = new CheckBox { Margin = new Thickness(20, 0, 0, 0) ,Height=32,Content="已完成"};
            addBtn = new Button { Margin = new Thickness(50, 0, 0, 0), Content = "+" };
            removeBtn = new Button { Margin = new Thickness(10, 0, 0, 0), Content = "-" };
            thisPanel.Children.Add(numLabel);
            thisPanel.Children.Add(comboBox);
            SetInputTip("請選擇下拉選單", projectTitle);
            thisPanel.Children.Add(projectTitle);
            SetInputTip("請輸入備註", projectDescription);
            thisPanel.Children.Add(projectDescription);
            SetInputTip("工時", hourSpent);
            thisPanel.Children.Add(hourSpent);
            thisPanel.Children.Add(doneCheck);
            addBtn.Click += AddButton_Clicked;
            thisPanel.Children.Add(addBtn);
            removeBtn.Click += RemoveButton_Clicked;
            thisPanel.Children.Add(removeBtn);
        }
        /// <summary>
        /// 設定這個欄位的編號，由主視窗顯示時一併設定
        /// </summary>
        /// <param name="num"></param>
        public void SetPanelNum(int num)
        {
            if(numLabel!=null)
            numLabel.Content=num.ToString();
        }
        /// <summary>
        /// 設定輸入格的提示
        /// </summary>
        /// <param name="inputTip"></param>
        /// <param name="target"></param>
        private void SetInputTip(string inputTip,TextBox target)
        {
            //創建VisualBrush作為TextBox的背景
            VisualBrush brush = new VisualBrush();
            TextBlock textBlock = new TextBlock
            {
                Text = inputTip,
                Foreground = new SolidColorBrush(Colors.Gray),
                FontSize = 15,
                FontStyle = FontStyles.Italic
            };
            brush.Visual = textBlock;
            brush.Stretch= Stretch.None;
            Style style = new Style(typeof(TextBox));
            //當文本為空時顯示佔位符
            style.Triggers.Add(new Trigger
            {
                Property = TextBox.TextProperty,
                Value = "",
                Setters = { new Setter(TextBox.BackgroundProperty, brush) }
            });
            //當TextBox獲得焦點時，清除背景
            style.Triggers.Add(new Trigger
            {
                Property = UIElement.IsFocusedProperty,
                Value=true,
                Setters = { new Setter(TextBox.BackgroundProperty,new SolidColorBrush(Colors.Transparent))}
            });
            //設置文本為空時背景顯示佔位符
            target.Background = brush;
            //當文本失去焦點且文本為空時顯示佔位符
            target.LostFocus += (sender, e) =>
            {
                TextBox? textbox = sender as TextBox;
                if (textbox != null && string.IsNullOrEmpty(textbox.Text))
                {
                    textbox.Background = brush;
                }
            };
            //當文本框獲得焦點時，清除背景使其透明
            target.GotFocus += (sender, e) =>
            {
                TextBox? textBox = sender as TextBox;
                textBox.Background = new SolidColorBrush(Colors.Transparent);
            };
            target.Style = style;
        }
        private void AddButton_Clicked(object sender,RoutedEventArgs e)
        {
            parentWindow.AddPanel();
        }
        private void RemoveButton_Clicked(object sender,RoutedEventArgs e)
        {
            parentWindow.RemovePanel(this);
        }

    }
}
