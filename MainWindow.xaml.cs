using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Word;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics;

namespace wordCounter
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ViewModel viewModel = new ViewModel();
        // 构造函数
        public MainWindow()
        {
            InitializeComponent();
            //窗口居中
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // 初始化viewmodel中的FileInfoList，view中也有对应的属性
            // ListView中的每一个item的数据结构（Id，FileName，Type，Size，Location，Status，Result,page ，word, line, paragraph）
            // 也跟FileInfo的数据结构相对应
            viewModel.FileInfoList = new ObservableCollection<FileInfo>();
            // datacontext绑定，这样view和viewmodel的变化就能相互更新
            this.FileList.DataContext = viewModel;
        }

        ////禁用右上角关闭按钮
        //protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        //{
        //    e.Cancel = true;
        //}

        // 添加文件的button的click对应的事件
        private void add_file_Click(object sender, RoutedEventArgs e)
        {
            var fileInfo = new FileInfo();
            // 添加文件对话框可对应的属性设置
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.InitialDirectory = @"C:\";
            // 文件类型过滤，就只能选择这些类型的文件
            dialog.Filter = "Word文档|*.docx;*.doc";//| Excel | *.xlsx | txt文件 | *.txt
            dialog.RestoreDirectory = true;
            //dialog.FilterIndex = 1;
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var file in dialog.FileNames)
                {
                    // 添加一个文件FileInfo对应的属性值
                    var id = viewModel.FileInfoList.Count + 1;//id
                    var sysFileInfo = new System.IO.FileInfo(file);
                    var fileName = sysFileInfo.Name;//文件名
                    var type = sysFileInfo.Extension;//格式
                    var size = sysFileInfo.Length / 1024; //大小 KB 
                    var location = sysFileInfo.DirectoryName;//位置
                    var status = "等待计算";//状态
                    var result = "";//字数
                    var page = "";//页数
                                  // var word = "";//字符数
                    var line = "";//行数
                    var paragraph = "";//段落

                    if (viewModel.FileInfoList.Any(a => a.FileName == fileName && a.Location == location))
                    {
                        System.Windows.MessageBox.Show("重复添加");
                        return;
                    }

                    // FileInfoList添加一个FileInfo，会自动更新到view
                    if ((type == ".xlsx" || type == ".docx" || type == ".doc" || type == ".txt" || page == ".xlsx" || page == ".docx") && !fileName.Contains("~$"))
                        viewModel.FileInfoList.Add(new FileInfo(id, fileName, type, size, location, status, result, page, line, paragraph));
                }

            }
        }

        // 添加文件夹的button的click对应的事件
        private void add_dir_Click(object sender, RoutedEventArgs e)
        {
            // 添加文件夹对话框
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.RootFolder = Environment.SpecialFolder.MyComputer;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 跟添加文件一样的，只不过添加的是对应文件下的所有文件
                // 其实以下代码可以提取出来重构出来一个函数，项目小，代码少就无所谓了
                string[] containFiles = System.IO.Directory.GetFiles(dialog.SelectedPath);
                for (int i = 0; i < containFiles.Length; i++)
                {
                    var id = viewModel.FileInfoList.Count + 1;
                    var sysFileInfo = new System.IO.FileInfo(containFiles[i]);
                    var fileName = sysFileInfo.Name;
                    var type = sysFileInfo.Extension;
                    var size = sysFileInfo.Length / 1024; // KB
                    var location = sysFileInfo.DirectoryName;
                    var status = "等待计算";
                    var result = "";
                    var page = "";
                    //  var word = "";
                    var line = "";
                    var paragraph = "";

                    if ((type == ".xlsx" || type == ".docx" || type == ".doc" || type == ".txt" || page == ".xlsx" || page == ".docx") && !fileName.Contains("~$"))
                        viewModel.FileInfoList.Add(new FileInfo(id, fileName, type, size, location, status, result, page, line, paragraph));

                }
            }
        }



        //开始计算的button的click对应的事件
        private void start_count_Click(object sender, RoutedEventArgs e)
        {
            //按下开始计算按钮后 button禁用
            start_count.IsEnabled = false;
            TextBlock1.Text = "0";
            TextBlock2.Text = "0";
            TextBlock3.Text = "0";
            TextBlock4.Text = "0";
            var list = new List<System.Threading.Tasks.Task>();
            //逐个计算FileInfoList中所有的文件的字数
            System.Threading.Tasks.Task.Run(() =>
            {
                foreach (var item in viewModel.FileInfoList)
                {
                    list.Add(System.Threading.Tasks.Task.Run(() =>//启动一个线程
                    {
                        if (item.Status != "计算完成")
                        {
                            // 实际的计算字数的函数，计算的结果也要更新到对应的item中去，view中相应的属性也会更新
                            item.Page = PageWord(System.IO.Path.Combine(item.Location, item.FileName), item.Type);
                            Dispatcher.Invoke(() => { TextBlock1.Text = (int.Parse(TextBlock1.Text) + int.Parse(item.Page)).ToString(); });
                            item.Result = CountWord(System.IO.Path.Combine(item.Location, item.FileName), item.Type);
                            Dispatcher.Invoke(() => { TextBlock2.Text = (int.Parse(TextBlock2.Text) + int.Parse(item.Result)).ToString(); });
                            //item.Word = WordWord(System.IO.Path.Combine(item.Location, item.FileName), item.Type);
                            item.Line = LineWord(System.IO.Path.Combine(item.Location, item.FileName), item.Type);
                            Dispatcher.Invoke(() => { TextBlock3.Text = (int.Parse(TextBlock3.Text) + int.Parse(item.Line)).ToString(); });
                            item.Paragraph = ParagraphWord(System.IO.Path.Combine(item.Location, item.FileName), item.Type);
                            Dispatcher.Invoke(() => { TextBlock4.Text = (int.Parse(TextBlock4.Text) + int.Parse(item.Paragraph)).ToString(); });
                            item.Status = "计算完成";

                        }
                    }));
                }
                // 等待所有线程结束
                System.Threading.Tasks.Task.WaitAll(list.ToArray());
                //重新启动button按钮
                Dispatcher.Invoke(() => { start_count.IsEnabled = true; });
            });
        }

        // 退出的button的click对应的事件
        private void exit_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            Process[] current = Process.GetProcesses();
            //遍历与当前进程名称相同的进程列表
            foreach (Process process in current)
            {
                //如果实例已经存在则kill当前进程
                if (process.ProcessName.ToUpper().Equals("WINWORD"))
                {
                    //杀死进程
                    process.Kill();
                    //等待程序关闭
                    process.WaitForExit();
                    //break;
                }
            }
            Close();
        }

        // 实际计算字数的函数
        private string CountWord(string filePath, string type)
        {
            // 根据文件类型的不同调用不同的函数计算
            if (type == ".docx")
                return CountDoc(filePath);
            else if (type == ".doc")
                return CountDoc(filePath);
            //else if (type == ".xlsx")
            //    return CountExcel(filePath);
            //else if (type == ".txt")
            //    return CountTxt(filePath);
            else
                return "不支持";
        }
        //计算word页数
        private string PageWord(string filePath, string page)
        {
            // 根据文件类型的不同调用不同的函数计算
            if (page == ".docx")
                return PageDoc(filePath);
            else if (page == ".doc")
                return PageDoc(filePath);
            else
                return "不支持";

        }

        ////计算word单词数
        //private string WordWord(string filePath, string word)
        //{
        //    // 根据文件类型的不同调用不同的函数计算
        //    if (word == ".docx")
        //        return WordDoc(filePath);
        //    else if (word == ".doc")
        //        return WordDoc(filePath);
        //    else
        //        return "不支持";
        //}

        //计算word文档行数
        private string LineWord(string filePath, string line)
        {
            // 根据文件类型的不同调用不同的函数计算
            if (line == ".docx")
                return LineDoc(filePath);
            else if (line == ".doc")
                return LineDoc(filePath);
            else
                return "不支持";
        }

        //计算word文档段落数
        private string ParagraphWord(string filePath, string paragraph)
        {
            // 根据文件类型的不同调用不同的函数计算
            if (paragraph == ".docx")
                return ParagraphDoc(filePath);
            else if (paragraph == ".doc")
                return ParagraphDoc(filePath);
            else
                return "不支持";
        }

        #region 使用Microsoft.Office.Interop.Word库来读取word文档、使用WdStatistic函数读取字数等
        // 计算word类型的文档的字数

        private string CountDoc(string filePath)
        {
            // 使用Microsoft.Office.Interop.Word库来读取word文档
            var wordApp = new MSWord.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            string ret = "";
            int count;

            try
            {
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(filePath);
                // 调用库的字数计算属性直接获取字数
                count = doc.ComputeStatistics(WdStatistic.wdStatisticCharacters, true);
                ret = count.ToString();

            }
            catch (Exception e)
            {
                ret = "计算出错";

            }
            finally
            {
                doc.Close();
            }

            return ret;
        }
        //计算word文档页数
        private string PageDoc(string filePath)
        {
            // 使用Microsoft.Office.Interop.Word库来读取word文档
            var wordApp = new MSWord.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            string ret = "";
            int pags;
            try
            {
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(filePath);
                // 调用库的字数计算属性直接获取页数

                pags = doc.ComputeStatistics(WdStatistic.wdStatisticPages, true);
                ret = pags.ToString();
            }
            catch (Exception e)
            {
                ret = "计算出错";
            }
            finally
            {
                doc.Close();
            }

            return ret;
        }

        

        //计算word文档行数
        private string LineDoc(string filePath)
        {
            // 使用Microsoft.Office.Interop.Word库来读取word文档
            var wordApp = new MSWord.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            string ret = "";
            int lines;
            try
            {
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(filePath);
                // 调用库的字数计算属性直接获取行数

                lines = doc.ComputeStatistics(WdStatistic.wdStatisticLines, true);
                ret = lines.ToString();
            }
            catch (Exception e)
            {
                ret = "计算出错";
            }
            finally
            {
                doc.Close();
            }

            return ret;
        }

        //计算word文档段落数
        private string ParagraphDoc(string filePath)
        {
            // 使用Microsoft.Office.Interop.Word库来读取word文档
            var wordApp = new MSWord.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            string ret = "";
            int paragraphs;
            try
            {
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(filePath);
                // 调用库的字数计算属性直接获取段落数

                paragraphs = doc.ComputeStatistics(WdStatistic.wdStatisticParagraphs, true);
                ret = paragraphs.ToString();
            }
            catch (Exception e)
            {
                ret = "计算出错";
            }
            finally
            {
                doc.Close();
            }

            return ret;
        }
        #endregion
        #region 计算word字单词数、Excel、TXT类型的字数
        //计算word文档单词数
        //private string WordDoc(string filePath)
        //{
        //    // 使用Microsoft.Office.Interop.Word库来读取word文档
        //    var wordApp = new MSWord.Application();
        //    Microsoft.Office.Interop.Word.Document doc = null;
        //    string ret = "";
        //    int words;
        //    try
        //    {
        //        wordApp.Visible = false;
        //        doc = wordApp.Documents.Open(filePath);
        //        // 调用库的字数计算属性直接获取单词数

        //        words = doc.ComputeStatistics(WdStatistic.wdStatisticWords, true);
        //        ret = words.ToString();
        //    }
        //    catch (Exception e)
        //    {
        //        ret = "计算出错";
        //    }
        //    finally
        //    {
        //        doc.Close();
        //    }

        //    return ret;
        //}

        ////计算excel类型的文档的字数
        //private string CountExcel(string filePath)
        //{
        //    // 使用Microsoft.Office.Interop.Excel库来读取Excel文档
        //    var excelApp = new MSExcel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook workbook = null;
        //    string ret = "";
        //    try
        //    {
        //        excelApp.Visible = false;
        //        workbook = excelApp.Workbooks.Open(filePath);
        //        int count = 0;
        //        // 选取的excel的所有的sheet都要循环计算并相加
        //        for (int i = 1; i <= workbook.Worksheets.Count; i++)
        //        {
        //            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[i];
        //            Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;

        //            // 每一个sheet的有数据的行列数
        //            var rowCount = ws.UsedRange.Rows.Count;
        //            var colCount = ws.UsedRange.Columns.Count;

        //            // 循环有数据的cell，计算每一个cell中的值的字数
        //            for (int m = 1; m <= rowCount; m++)
        //            {
        //                for (int n = 1; n <= colCount; n++)
        //                {
        //                    // m, n used carefully
        //                    var temp = ws.Cells[n][m];
        //                    if (temp.Value != null)
        //                        count = count + temp.Value.ToString().Length;
        //                }
        //            }
        //        }
        //        ret = count.ToString();
        //    }
        //    catch (Exception e)
        //    {
        //        ret = "计算出错";
        //    }
        //    finally
        //    {
        //        workbook.Close();
        //    }
        //    return ret;
        //}

        //// 计算txt类型的文档的字数
        //// txt类型的文档比较直接全部读取计算长度即可
        //private string CountTxt(string filePath)
        //{
        //    var ret = "";
        //    try
        //    {
        //        StreamReader sr = new StreamReader(filePath);
        //        ret = sr.ReadToEnd().Length.ToString();
        //    }
        //    catch (Exception)
        //    {
        //        ret = "计算出错";
        //    }
        //    return ret;
        //}

        #endregion

        // 右键删除函数button的click对应的事件
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 调用DelItemInFileInfoList函数删除对应项，参数是选中的item的index
            DelItemInFileInfoList(this.FileList.SelectedIndex);
        }

        private void DelItemInFileInfoList(int selectedIndex)
        {
            // 对应的index，把FileInfoList中的对应项删除即可
            // 此时要更新比index大的Id，都减一
            //if (viewModel.FileInfoList.Count == 0) return;
            if (viewModel.FileInfoList.Count == 0 || selectedIndex == -1) return;
            viewModel.FileInfoList.RemoveAt(selectedIndex);
            foreach (var item in viewModel.FileInfoList)
            {
                if (item.Id > 0 && item.Id > selectedIndex)
                {
                    item.Id = item.Id - 1;
                }
            }
        }

        // 右键全部移除的button的click对应的事件
        private void rm_all_Click(object sender, RoutedEventArgs e)
        {
            // 将viewmodel的FileInfoList清空，view也就清空了
            viewModel.FileInfoList.Clear();
        }


        //导出至Excel
        private void export_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                //命名为时间
                FileName = DateTime.Now.ToString("yyyyMMddHHmmss"),
                Filter = "Excel 工作簿|*.xlsx",
            };
            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            export.IsEnabled = false;
            IWorkbook workbook = new XSSFWorkbook();//创建工作簿
            ISheet sheet = workbook.CreateSheet();//创建工作表
            {
                IRow row = sheet.CreateRow(sheet.LastRowNum);//创建一行
                var columns = (FileList.View as GridView).Columns;//获取所有列名
                for (int i = 0; i < columns.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(columns[i].Header.ToString());//往第i个单元格里写入数据
                }
            }
            foreach (var item in viewModel.FileInfoList)
            {
                IRow row = sheet.CreateRow(sheet.LastRowNum + 1);
                row.CreateCell(0).SetCellValue(item.Id);
                row.CreateCell(1).SetCellValue(item.FileName);
                row.CreateCell(2).SetCellValue(item.Type);
                row.CreateCell(3).SetCellValue(item.Size);
                row.CreateCell(4).SetCellValue(item.Location);
                row.CreateCell(5).SetCellValue(item.Status);
                row.CreateCell(6).SetCellValue(item.Page);
                row.CreateCell(7).SetCellValue(item.Result);
                //row.CreateCell(8).SetCellValue(item.Word);
                row.CreateCell(8).SetCellValue(item.Line);
                row.CreateCell(9).SetCellValue(item.Paragraph);
            }
            using (FileStream stream = File.Create(dialog.FileName))
            {
                workbook.Write(stream);//写入并保存文件
            }
            export.IsEnabled = true;
        }

        //点击列头排序功能
        private ListSortDirection _sortDirection;
        private GridViewColumnHeader _sortColumn;
        private string processName;
        private readonly object lstPostion;

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            GridViewColumnHeader column = e.OriginalSource as GridViewColumnHeader;
            if (column == null || column.Column == null)
            {
                return;
            }
            if (_sortColumn == column)
            {
                // 切换排序方向 
                _sortDirection = _sortDirection == ListSortDirection.Ascending ? ListSortDirection.Descending : ListSortDirection.Ascending;
            }
            else
            {
                // 从以前排序的标题中删除箭头 
                if (_sortColumn != null && _sortColumn.Column != null)
                {
                    _sortColumn.Column.HeaderTemplate = null;
                    _sortColumn.Column.Width = _sortColumn.ActualWidth - 20;
                }
                _sortColumn = column;
                _sortDirection = ListSortDirection.Ascending;
                column.Column.Width = column.ActualWidth + 20;
            }
            if (_sortDirection == ListSortDirection.Ascending)
            {
                column.Column.HeaderTemplate = Resources["ArrowUp"] as DataTemplate;
            }
            else
            {
                column.Column.HeaderTemplate = Resources["ArrowDown"] as DataTemplate;
            }
            string header = string.Empty;

            // 如果使用绑定且属性名与头内容不匹配
            System.Windows.Data.Binding b = _sortColumn.Column.DisplayMemberBinding as System.Windows.Data.Binding;
            if (b != null)
            {
                header = b.Path.Path;
            }

            ICollectionView resultDataView = CollectionViewSource.GetDefaultView((sender as System.Windows.Controls.ListView).ItemsSource);
            resultDataView.SortDescriptions.Clear();
            resultDataView.SortDescriptions.Add(new SortDescription(header, _sortDirection));
        }
    }


    

    // FileInfo数据接口定义，跟view中的ListView的item数据结构对应
    public class FileInfo : INotifyPropertyChanged
    {
        public int _id;
        public string _fileName;
        public string _type;
        public double _size;
        public string _location;
        public string _status;
        public string _result;
        public string _page;
        // public string _word;
        public string _line;
        public string _paragraph;
        // INotifyPropertyChanged的接口实现，能和view相互更新
        public event PropertyChangedEventHandler PropertyChanged;

        public int Id
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
                // INotifyPropertyChanged的接口实现，能和view相互更新，set即赋值的时候可以更新到view;每个属性都一样做这个实现

                if (this.PropertyChanged != null)//激发事件
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Id"));
                }
            }
        }

        public string FileName
        {
            get
            {
                return _fileName;
            }
            set
            {
                _fileName = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("FileName"));
                }
            }
        }


        public string Type
        {
            get
            {
                return _type;
            }
            set
            {
                _type = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Type"));
                }
            }
        }

        public double Size
        {
            get
            {
                return _size;
            }
            set
            {
                _size = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Size"));
                }
            }
        }

        public string Location
        {
            get
            {
                return _location;
            }
            set
            {
                _location = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Location"));
                }
            }
        }

        public string Status
        {
            get
            {
                return _status;
            }
            set
            {
                _status = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Status"));
                }
            }
        }

        public string Result
        {
            get
            {
                return _result;
            }
            set
            {
                _result = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Result"));
                }
            }
        }

        public string Page
        {
            get
            {
                return _page;
            }
            set
            {
                _page = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Page"));
                }
            }
        }

        //public string Word
        //{
        //    get
        //    {
        //        return _word;
        //    }
        //    set
        //    {
        //        _word = value;
        //        if (this.PropertyChanged != null)
        //        {
        //            this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Word"));
        //        }
        //    }
        //}

        public string Line
        {
            get
            {
                return _line;
            }
            set
            {
                _line = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Line"));
                }
            }
        }

        public string Paragraph
        {
            get
            {
                return _paragraph;
            }
            set
            {
                _paragraph = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Paragraph"));
                }
            }
        }

        public FileInfo() { }
        public FileInfo(int id, string fileName, string type, double size, string location, string status, string result, string page, string line, string paragraph)
        {
            _id = id;
            _fileName = fileName;
            _type = type;
            _size = size;
            _location = location;
            _status = status;
            _result = result;
            _page = page;
            // _word = word;
            _line = line;
            _paragraph = paragraph;
        }

        public FileInfo(int id, string fileName, string type, long size, string location, string status, string result, string page, string line, string paragraph)
        {
            Id = id;
            FileName = fileName;
            Type = type;
            Size = size;
            Location = location;
            Status = status;
            Result = result;
            Page = page;
            //  Word = word;
            Line = line;
            Paragraph = paragraph;
        }
    }

    // viewmodel定义，其中的FileInfoList跟view上的FileList对应
    // 两者绑定，相互更新
    public class ViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<FileInfo> _fileInfoList;
        public ObservableCollection<FileInfo> FileInfoList
        {
            get
            {
                return this._fileInfoList;
            }
            set
            {
                if (this._fileInfoList != value)
                {
                    this._fileInfoList = value;
                    OnPropertyChanged("FileInfoList");
                }
            }
        }

        // INotifyPropertyChanged的接口实现，能和view相互更新
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
