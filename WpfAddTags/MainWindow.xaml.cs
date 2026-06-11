using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace WpfAddTags
{
    
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<TagInfo> GeneralTagsList;
        string GeneralFilePath = string.Empty;
        string KepToTable = string.Empty;
        public ObservableCollection<TagInfo> ExeptTags { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            lstTags.SelectionChanged += LstTags_SelectionChanged;
        }

        private void LstTags_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectTagsCount.Text=lstTags.SelectedItems.Count.ToString();
        }
        #region Buttons
        private void ButtonGeneralFile_Click(object sender, RoutedEventArgs e)
        {
                GeneralFilePath = ChooseFile();
                LabelGeneralFile.Text = GeneralFilePath;
            
        }
        private void ButtonKepToTableFile_Click(object sender, RoutedEventArgs e)
        {
                KepToTable = ChooseFile();
                LabelKepToTableFile.Text = KepToTable;  
        }
        private void ButtonCompareFiles_Click(object sender, RoutedEventArgs e)
        {
            if (GeneralFilePath == string.Empty || KepToTable == string.Empty) { MessageBox.Show("Не выбраны файлы!"); return; }
            GeneralTagsList = ExcelRead(GeneralFilePath, 0, "A1:P");
            GeneralTagsList = GeneralTagsList.OrderBy(t => t.NumTag).ToList();
            var NewTagsList = ExcelRead(KepToTable, 4, "A1:T");
            var ExceptList = NewTagsList.Except(GeneralTagsList); //все теги, которых нет в старом списке

            if (ExceptList.Count() == 0) { MessageBox.Show("Нет новых тегов"); return; }
            lstTags.ItemsSource = ExceptList.ToList();
            lstTags.DisplayMemberPath = "Item";
            newTagsCount.Text = ExceptList.Count().ToString();
            CheckCountTags(GeneralTagsList, NewTagsList, ExceptList.ToList());
        }

        private void ButtonSelectAll_Click(object sender, RoutedEventArgs e)
        {
            if (lstTags.SelectedIndex != -1) { lstTags.SelectedIndex = -1; }
            else
                lstTags.SelectAll();

        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            WriteExcel(GeneralFilePath, lstTags.SelectedItems);
        }


        private void ButtonCheckDuble_Click(object sender, RoutedEventArgs e)
        {
            Cheking.Visibility = Visibility.Visible;
            string filename = ChooseFile();
            if (filename != string.Empty)
            {
                var checklist = ExcelRead(filename, 0, "A1:P");

                bool haveItemDuble = false;
                bool haveNameDuble = false;
                bool haveNumTagDuble = false;
                foreach (var taginfo in checklist)
                {
                    var chekItemDuble = checklist.Where(i => i.Item == taginfo.Item);
                    if (chekItemDuble.Count() > 1)
                    {
                        haveItemDuble = chekItemDuble != null;
                        WriteTxt("Item: "+taginfo.Item);
                    }

                    var chekNameDuble = checklist.Where(i => i.Name == taginfo.Name);
                    if (chekNameDuble.Count() > 1)
                    {
                        haveNameDuble = chekNameDuble != null;
                        WriteTxt("Name: " + taginfo.Name);

                    }

                    var chekNumTagDuble = checklist.Where(i => i.NumTag == taginfo.NumTag);
                    if (chekNumTagDuble.Count() > 1)
                    {
                        haveNumTagDuble = chekNumTagDuble != null;
                        WriteTxt("NumTag: " + taginfo.NumTag);
                    }

                }
                if (haveItemDuble) MessageBox.Show("Повторяются значения в первом столбце");
                if (haveNameDuble) MessageBox.Show("Повторяются значения во втором столбце");
                if (haveNumTagDuble) MessageBox.Show("Неправильная нумерация");

                if (haveItemDuble && haveNameDuble && haveNumTagDuble) Process.Start(Directory.GetCurrentDirectory()+"\\log.txt");
                else MessageBox.Show("Все в порядке");
            }
            Cheking.Visibility = Visibility.Collapsed;

        }
        #endregion
        string ChooseFile()
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Text documents (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.FilterIndex = 2;

            Nullable<bool> result = dialog.ShowDialog();

            if (result == true)
            {
                return dialog.FileName;
            }
            return string.Empty;
        }
        /// <summary>
        /// чтение excel
        /// </summary>
        /// <param name="name">полное имя файла</param>
        /// <param name="firstColumn">колонка с которой начинается информация</param>
        /// <param name="lastColumn">последняя колонка с информацияей</param>
        /// <returns></returns>
        List<TagInfo> ExcelRead(string name, int firstColumn, string lastColumn)
        {
            //считываем данные из Excel файла в двумерный массив
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlSht; //лист Excel   
            xlWB = xlApp.Workbooks.Open(name); //название файла Excel                                             
            xlSht = xlWB.Worksheets[1]; //1-й лист в книге
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            var arrData = (object[,])xlSht.Range[lastColumn + iLastRow].Value; //берём данные с листа Excel
            //xlApp.Visible = true; //отображаем Excel     
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel

            int RowsCount = arrData.GetUpperBound(0);
            int ColumnsCount = arrData.GetUpperBound(1);
            List<TagInfo> list = new List<TagInfo>();

            //заполняем list данными из массива
            int j;
            for (j = 1; j <= RowsCount - 1; j++)
            {
                if (arrData[j + 1, firstColumn + 1] == null) break;
                list.Add(new TagInfo(arrData, j + 1, firstColumn));

            }
            return list;
        }
        void WriteExcel(string fileName,IList tags)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            var lastnum = GeneralTagsList.Last().NumTag;
            for (int i = 0; i < tags.Count; i++)
            {
                var tag=(TagInfo)tags[i];
                iLastRow++;
                xlSht.Cells[iLastRow, "A"].Value = tag.Item;
                xlSht.Cells[iLastRow, "B"].Value = tag.Name;
                xlSht.Cells[iLastRow, "C"].Value = tag.Type;
                xlSht.Cells[iLastRow, "D"].Value = tag.TypeRW;
                xlSht.Cells[iLastRow, "E"].Value = tag.AWU;
                xlSht.Cells[iLastRow, "F"].Value = tag.Legend;
                xlSht.Cells[iLastRow, "G"].Value = tag.NumStation;
                xlSht.Cells[iLastRow, "H"].Value = lastnum+1;
                xlSht.Cells[iLastRow, "I"].Value = tag.CAAG;
                xlSht.Cells[iLastRow, "J"].Value = tag.Koef;
                xlSht.Cells[iLastRow, "K"].Value = tag.Color;
                xlSht.Cells[iLastRow, "L"].Value = tag.Init;
                xlSht.Cells[iLastRow, "M"].Value = tag.Group;
                xlSht.Cells[iLastRow, "N"].Value = tag.ChangeVal;
                xlSht.Cells[iLastRow, "O"].Value = tag.SaveByTime;
                xlSht.Cells[iLastRow, "P"].Value = tag.RarelyChanging;
                lastnum++;
            }
            var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Workbook|*.xlsx";
                saveFileDialog.Title = "Сохранить файл Excel";
                saveFileDialog.FileName = ""; // Имя по умолчанию

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Получаем выбранный путь и сохраняем файл
                    string filePath = saveFileDialog.FileName;
                    xlWb.SaveAs(filePath);

                    MessageBox.Show("Файл успешно сохранен!");
                }
            
            //xlApp.Visible = true;
            xlWb.Close(false); //закрыть и не сохранить книгу
            xlApp.Quit();
            var sum = GeneralTagsList.Count + tags.Count;
            if (sum != lastnum) MessageBox.Show("Проверьте конечный файл, возможна ошибка: неправильная нумерация");

        }

        void WriteTxt(string info)
        {
            string path = "log.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(info);
            }
        }
        void CheckCountTags(List<TagInfo> generaltaglist, List<TagInfo> newtaglist, List<TagInfo> exeptlist )
        {
            var sum=generaltaglist.Count+exeptlist.Count;
            if (sum != newtaglist.Count) MessageBox.Show("Проверьте конечный файл, возможна ошибка: повтор тегов");
        }

    }
}
