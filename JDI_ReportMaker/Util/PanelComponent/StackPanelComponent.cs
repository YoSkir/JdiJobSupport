using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace JDI_ReportMaker.Util.PanelComponent
{
    public abstract class StackPanelComponent
    {
        public abstract StackPanel GetPanel();
        public abstract StackPanel CreatePanel();

        public abstract void AddPanel();
        public abstract void RemovePanel();




        /// <summary>
        /// 設定輸入格的提示
        /// </summary>
        /// <param name="inputTip"></param>
        /// <param name="target"></param>
        internal void SetInputTip(string inputTip, TextBox target)
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
            brush.Stretch = Stretch.None;
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
                Value = true,
                Setters = { new Setter(TextBox.BackgroundProperty, new SolidColorBrush(Colors.Transparent)) }
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
    }
}
