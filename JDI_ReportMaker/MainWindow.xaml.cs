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
        public MainWindow()
        {
            InitializeComponent();
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string[] filePath = filePathTextBox.Text.Split('.');
            try
            {
                createTestXls(filePathTextBox.Text, filePath[0]+"幹." + filePath[1]);
                //CopyExcelFile(filePathTextBox.Text, filePath[0]+" test ." + filePath[1]);
                resultLabel.Content = "儲存成功";
            }
            catch(Exception ex)
            {
                resultLabel.Content = "儲存失敗";
                testyo.Text= ex.Message;
                MessageBox.Show(ex.Message);
            }

        }
        private void createTestXls(string sourceFilePath,string targetFilePath)
        {
            try
            {
                IWorkbook sourceWorkbook;
                using (FileStream file=new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    sourceWorkbook=new HSSFWorkbook(file);
                }
                using(FileStream file=new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                {
                    sourceWorkbook.Write(file);
                }

            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 複製並另存xls檔
        /// </summary>
        /// <param name="sourceFilePath">原檔的位置與名稱</param>
        /// <param name="targerFilePath">另存檔案的位置與名稱</param>
        private void CopyExcelFile(string sourceFilePath,string targerFilePath)
        {
            try
            {
                using (FileStream sourceFile = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook sourceWorkbook;
                    if (System.IO.Path.GetExtension(sourceFilePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        sourceWorkbook = new HSSFWorkbook(sourceFile);
                    }
                    else
                    {
                        MessageBox.Show("檔案格式不正確，請選取.xls檔");
                        return;
                    }

                    IWorkbook targetWorkbook = new HSSFWorkbook();
                    //遍歷工作表
                    for (int i = 0; i < sourceWorkbook.NumberOfSheets; i++)
                    {
                        ISheet sourceSheet = sourceWorkbook.GetSheetAt(i);
                        ISheet targetSheet = targetWorkbook.CreateSheet(sourceSheet.SheetName);
                        //遍歷工作表的行
                        for (int j = 0; j < sourceSheet.LastRowNum; j++)
                        {
                            IRow sourceRow = sourceSheet.GetRow(j);
                            if(sourceRow != null)
                            {
                                IRow targetRow = targetSheet.CreateRow(j);
                                //遍歷工作表的格
                                for (int k = 0; k < sourceRow.LastCellNum+1; k++)
                                {
                                    ICell sourceCell = sourceRow.GetCell(k);
                                    if (sourceCell != null)
                                    {
                                        ICell targetCell = targetRow.CreateCell(k, sourceCell.CellType);
                                        switch (sourceCell.CellType)
                                        {
                                            case CellType.Boolean:
                                                targetCell.SetCellValue(sourceCell.BooleanCellValue); break;
                                            case CellType.String:
                                                targetCell.SetCellValue(sourceCell.StringCellValue); break;
                                            case CellType.Numeric:
                                                targetCell.SetCellValue(sourceCell.NumericCellValue);break;
                                            default:targetCell.SetCellValue(sourceCell.RichStringCellValue);break ;
                                        }
                                        ICellStyle cellStyle = targetWorkbook.CreateCellStyle();
                                        cellStyle.CloneStyleFrom(sourceCell.CellStyle);
                                        targetCell.CellStyle = cellStyle;
                                    }
                                }
                            }
                            mergeRegion(sourceSheet, targetSheet);
                        }
                    }
                    //另存新檔
                    using(FileStream newFile=new FileStream(targerFilePath, FileMode.Create, FileAccess.Write))
                    {
                        targetWorkbook.Write(newFile);
                    }
                }
            }catch
            {
                throw ;
            }
        }
        private void mergeRegion(ISheet sourceSheet,ISheet targetSheet)
        {
            int numberOfMergedRegion = sourceSheet.NumMergedRegions;
            for(int mergeIndex=0; mergeIndex<numberOfMergedRegion; mergeIndex++)
            {
                CellRangeAddress mergedRegion=sourceSheet.GetMergedRegion(mergeIndex);
                try
                {
                    targetSheet.AddMergedRegion(new CellRangeAddress(mergedRegion.FirstRow, mergedRegion.LastRow,
                                                                    mergedRegion.FirstColumn, mergedRegion.LastColumn));
                }
                catch { }
            }
        }
    }
}