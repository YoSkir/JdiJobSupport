using MahApps.Metro.Controls;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace JDI_ReportMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private readonly Dictionary<string, string> fileName 
            = new Dictionary<string, string> { {"日報表", "佳帝科技員工日報表 "},
                                               {"周報表", "佳帝科技員工周報表 "},
                                               {"工時表", "佳帝科技員工工時表 "}};
        public MainWindow()
        {
            initialData();
            InitializeComponent();
        }
        /// <summary>
        /// 初始化資料
        /// </summary>
        private void initialData()
        {
        }

        private void ButtonSelectExcel_Click(Object sender,RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel File|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath=openFileDialog.FileName;
                filePathTextBox.Text= filePath;
            }
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (inputCheck())
            {
                try
                {
                    //做報表種類判斷
                    IWorkbook target = loadFile(filePathTextBox.Text);
                    staffDataWrite(target);
                    saveFile(target, FileNameEnum.日報表);
                    resultLabel.Content = "儲存成功";
                }catch (Exception ex)
                {
                    resultLabel.Content = "儲存失敗";
                    logLabel.Content = ex.Message;
                }

            }
            else
            {
                logLabel.Content = "請確認檔案路徑、員工資料";
            }
        }
        /// <summary>
        /// 確認員工姓名、編號、來源檔案、目的地是否有輸入
        /// </summary>
        /// <returns></returns>
        private bool inputCheck()
        {
            return nameTextbox.Text.Length > 0&&
                departmentCbox.Text.Length > 0&&
                filePathTextBox.Text.Length > 0&&
                savePathTextBox.Text.Length>0;
        }
        /// <summary>
        /// 將員工姓名、日期寫入
        /// </summary>
        private void staffDataWrite(IWorkbook target)
        {
            overwriteFile(target, 0, 5, 0, "姓名: "+nameTextbox.Text);
            overwriteFile(target, 0, 5, 3, "部門: "+departmentCbox.Text);
            overwriteFile(target, 0, 5, 7, "日期: "+DateTime.Now.ToString("yyyy/MM/dd"));
        }
        /// <summary>
        /// 檔案資料單格寫入(不取代原有內容)
        /// </summary>
        /// <param name="target">目標檔案</param>
        /// <param name="sheet">工作表</param>
        /// <param name="row">列</param>
        /// <param name="cell">格</param>
        /// <param name="value">指定內容</param>
        private void writeFile(IWorkbook target,int sheet,int row,int cell,string value)
        {
            try
            {
                ISheet isheet = target.GetSheetAt(sheet);
                if(isheet != null)
                {
                    IRow irow=isheet.GetRow(row);
                    if (irow != null)
                    {
                        ICell icell=irow.GetCell(cell);
                        if(icell != null)
                        {
                            icell.SetCellValue(icell.StringCellValue+" "+value);
                        }
                    }
                }
            }
            catch
            {
                resultLabel.Content = "寫入失敗";
                throw;
            }
        }
        /// <summary>
        /// 檔案資料單格寫入(取代原有內容)
        /// </summary>
        /// <param name="target">目標檔案</param>
        /// <param name="sheet">工作表</param>
        /// <param name="row">列</param>
        /// <param name="cell">格</param>
        /// <param name="value">指定內容</param>
        private void overwriteFile(IWorkbook target, int sheet, int row, int cell, string value)
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
                resultLabel.Content = "寫入失敗";
                throw;
            }
        }

        /// <summary>
        /// 讀取目標報表
        /// </summary>
        /// <param name="sourcePath">目標報表之路徑</param>
        /// <returns>回傳讀取完之報表</returns>
        private IWorkbook loadFile(string sourcePath)
        {
            IWorkbook sourceWorkbook;
            try
            {
                using(FileStream file=new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
                {
                    sourceWorkbook = new HSSFWorkbook(file);
                }
            }
            catch
            {
                resultLabel.Content = "讀取失敗";
                throw;
            }
            return sourceWorkbook;
        }
        /// <summary>
        /// 另存xls檔
        /// </summary>
        /// <param name="sourceWorkbook">原檔</param>
        /// <param name="targerFilePath">另存檔案的位置與名稱</param>
        private void saveFile(IWorkbook sourceWorkbook,FileNameEnum fileNameEnum)
        {
            //設定檔名:報表名+員工名+日期
            string? date = datePicker.Text.Length == 0 ? DateTime.Today.ToString("MMdd") :
                datePicker.SelectedDate?.ToString("MMdd");
            string fileNameStr = "/" + fileName[fileNameEnum.ToString()] + " " + nameTextbox.Text + " " + date + ".xls";
            string targetFilePath=savePathTextBox.Text+ fileNameStr;
            try
            {
                using(FileStream file=new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                {
                    sourceWorkbook.Write(file);
                }
            }
            catch
            {
                resultLabel.Content = "儲存失敗";
                throw;
            }
        }


        private void dataConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if(nameTextbox.Text.Length < 1)
            {
                logLabel.Content="請填寫員工姓名";
            }
            else if(departmentCbox.Text.Length < 1)
            {
                logLabel.Content = "請選擇部門";
            }
            else
            {
                string date = DateTime.Now.ToString("yyyy/MM/dd");
                if(datePicker.Text.Length > 0)
                {
                    date=datePicker.Text;
                }
                testyo.Text = nameTextbox.Text + "\n" + departmentCbox.Text+"\n"+date;
            }
        }

        private void selectLocateButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFolderDialog openFolderDialog = new OpenFolderDialog() {
                Title = "請選擇目標位置",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                Multiselect = false
            } ;
            if (openFolderDialog.ShowDialog() == true)
            {
                savePathTextBox.Text = openFolderDialog.FolderName;
            }
        }
    }
}