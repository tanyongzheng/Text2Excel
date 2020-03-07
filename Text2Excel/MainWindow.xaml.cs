using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using MahApps.Metro.Controls;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Text2Excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {

        private DataTable dataTable;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_SourceFilePath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "文本文件|*.txt";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "txt";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_SourceFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btn_TargetDirPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_TargetDirPath.Text = openFileDialog.SelectedPath;
            }
        }

        private async void btn_ConvertFile_Click(object sender, RoutedEventArgs e)
        {
            string sourceFilePath = txt_SourceFilePath.Text;
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                MessageBox.Show("请选择要转换的文件！");
                return;
            }
            string targetDire = txt_TargetDirPath.Text;
            if (string.IsNullOrEmpty(targetDire))
            {
                MessageBox.Show("请选择目标文件夹！");
                return;
            }

            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("转换的源文件不存在！");
                return;
            }
            if (!Directory.Exists(targetDire))
            {
                MessageBox.Show("目标文件夹不存在！");
                return;
            }
            var columnSplitStr = this.txt_ColumnSplitStr.Text;
            if (string.IsNullOrEmpty(columnSplitStr))
            {
                MessageBox.Show("列分隔符不能为空！");
                return;
            }
            loadingBar.IsOpen = true;
            txt_Msg.Text = "";
            dataTable = new DataTable();
            var rowList = GetFileRows(sourceFilePath);
            await foreach (var row in rowList)
            {

            }
            var indexColumnName= txt_IndexColumnName.Text;
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(indexColumnName))
                {
                    dataTable.Columns.Add(indexColumnName, typeof(int));
                    for(var i = 0; i < dataTable.Rows.Count; i++)
                    {
                        dataTable.Rows[i][indexColumnName] = i + 1;
                    }
                }
            }
            if (string.IsNullOrEmpty(txt_Msg.Text))
            {
                txt_Msg.Text = $"总共{dataTable.Rows.Count}行";
                var excelFileName = System.IO.Path.GetFileNameWithoutExtension(sourceFilePath) + ".xlsx";
                var saveExcelPath = targetDire+"\\" + excelFileName;
                //var exportFileInfo =
                await Task.Run(async () =>
                {
                    IExporter exporter = new ExcelExporter();
                    await exporter.Export(saveExcelPath, dataTable);
                    MessageBox.Show("转换成功！");
                });
            }
            loadingBar.IsOpen = false;
        }

        private async IAsyncEnumerable<DataRow> GetFileRows(string filePath)
        {
            StreamReader reader = File.OpenText(filePath);
            var index = 0;
            var columnCount = 0;
            var columnNames = new List<string>();
            while (true)
            {
                var currentStr = await reader.ReadLineAsync();
                if (string.IsNullOrEmpty(currentStr))
                {
                    //txt_Msg.Text = $"第{index+1}行字为空！";
                    break;
                }
                if (!currentStr.Contains(txt_ColumnSplitStr.Text))
                {
                    txt_Msg.Text = $"第{index+1}行没有分割符{txt_ColumnSplitStr.Text}！";
                    break;
                }
                var lineStrArray = currentStr.Split(txt_ColumnSplitStr.Text);
                if (index == 0)
                {
                    columnCount = lineStrArray.Length;
                    if (lineStrArray.Length != lineStrArray.Distinct().Count())
                    {
                        txt_Msg.Text = $"首行有重复列名！";
                        break;
                    }
                    foreach (var columnName in lineStrArray)
                    {
                        dataTable.Columns.Add(columnName, typeof(string));
                    }
                    columnNames.AddRange(lineStrArray);
                    index++;
                }
                else
                {
                    if (lineStrArray.Length != columnCount)
                    {
                        txt_Msg.Text = $"第{index + 1}行分割的列数（{lineStrArray.Length}）和首行列数（{columnCount}）不一致，无法转换！";
                        break;
                    }

                    var dataRow = dataTable.NewRow();
                    for(var i=0;i<columnNames.Count;i++)
                    {
                        var columnName = columnNames[i];
                        dataRow[columnName] = lineStrArray[i];
                    }
                    dataTable.Rows.Add(dataRow);
                    index++;
                    yield return dataRow;
                }
            }
            reader.Close();
            reader.Dispose();
        }

    }
}
