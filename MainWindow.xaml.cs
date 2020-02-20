using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ExcelAppWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application _excel = null;
        private Excel.Workbook _refWb = null;
        private Excel.Workbook _wb = null;
        private string _workBookFileExtension = null;
        public MainWindow()
        {
            _excel = Application.Current.Resources["Excel"] as Excel.Application;
            InitializeComponent();
        }
        private string fileExtension = null;
        private void btnLoadReferenceFile(object sender, RoutedEventArgs e)
        {
            try {
                _refWb?.Close(false);
            } catch { }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel-compatible files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                ReferenceFile.Text = openFileDialog.FileName;
            }
        }
        private void ReferenceFile_Drop(object sender, DragEventArgs e)
        {
            try { _refWb?.Close(false); }
            catch { }
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0)
                {
                    ReferenceFile.Text = files[0];
                }
            }
        }
        private void ReferenceFile_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }


        private void btnLoadFileToCheck(object sender, RoutedEventArgs e)
        {
            try
            {
                _wb?.Close(false);
            } catch { }
            Result.Clear();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel-compatible files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                FileToCheck.Text = openFileDialog.FileName;
                string[] filenameSplits = openFileDialog.FileName.Split('.');
                _workBookFileExtension = filenameSplits[filenameSplits.Length - 1];
                BtnRunProgram.IsEnabled = true;
            }
        }
        private void FileToCheck_Drop(object sender, DragEventArgs e)
        {
            try
            {
                _wb?.Close(false);
            } catch { }
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0)
                {
                    FileToCheck.Text = files[0];
                    string[] filenameSplits = files[0].Split('.');
                    _workBookFileExtension = filenameSplits[filenameSplits.Length - 1];
                    BtnRunProgram.IsEnabled = true;
                }
            }
        }
        private void FileToCheck_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void runProgram(object sender, EventArgs e)
        {
            Result.Clear();
            Result.AppendText("Working...");
            try
            {
                _refWb = _excel.Workbooks.Open(ReferenceFile.Text, false, true);
                _wb = _excel.Workbooks.Open(FileToCheck.Text, false, true);
                List<string> ancestorWorksheetNames = new List<string>();
                foreach (Excel.Worksheet ws in _refWb.Worksheets)
                {
                    ancestorWorksheetNames.Add(ws.Name);
                }

                foreach (Excel.Worksheet ws in _wb.Worksheets)
                {
                    if (!ancestorWorksheetNames.Contains(ws.Name))
                    {
                        Result.AppendText(Environment.NewLine + $"new worksheet found: {ws.Name}");
                        ws.Cells.Interior.Color = ConvertColor(System.Drawing.Color.Yellow);
                        continue;
                    }
                    var rng = ws.UsedRange;
                    for (int i = 1; i <= rng.Rows.Count; i++)
                    {
                        for (int j = 1; j <= rng.Columns.Count; j++)
                        {
                            var ancestorWs = _refWb.Worksheets[ws.Name];
                            var ancestorCell = ancestorWs.Cells[i, j];
                            bool sameCell = false;
                            try
                            {
                                sameCell = rng.Cells[i, j].Value == ancestorCell.Value;
                            }
                            catch
                            {
                            }
                            finally
                            {
                                if (!sameCell)
                                {
                                    Result.AppendText(Environment.NewLine + $"revision found at {ws.Name}, row {i}, col {j}: {rng.Cells[i, j].Value.ToString()}");
                                    rng.Cells[i, j].Interior.Color = ConvertColor(System.Drawing.Color.Yellow);
                                    if (ancestorCell.Value == null)
                                        rng.Cells[i, j].AddComment("old value is empty/null");
                                    else
                                    {
                                        rng.Cells[i, j].AddComment("old value:" + ancestorCell.Value.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                BtnSaveFile.IsEnabled = true;
            }
            catch (Exception ex)
            {
                Result.AppendText(Environment.NewLine + $"ERROR: " + ex.ToString());
                _wb?.Close(false);
            }
            finally
            {
                _refWb?.Close(false);
            }
        }
        private void saveFile(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = $"Excel file (*.{_workBookFileExtension})|*.{_workBookFileExtension}";
            if (saveFileDialog.ShowDialog() == true)
            {
                _wb?.SaveAs(saveFileDialog.FileName);
                Result.AppendText(Environment.NewLine + "File saved as: " + saveFileDialog.FileName);
                _wb?.Close(false);
                BtnSaveFile.IsEnabled = false;
            }
        }
        public static int ConvertColor(System.Drawing.Color color)
        {
            int r = color.R;
            int g = color.G * 256;
            int b = color.B * 65536;
            return r + g + b;
        }
    }
}
