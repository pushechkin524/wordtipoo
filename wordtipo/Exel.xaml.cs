using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
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
using System.Windows.Shapes;
using static MaterialDesignThemes.Wpf.Theme;

namespace wordtipo
{
    /// <summary>
    /// Логика взаимодействия для Exel.xaml
    /// </summary>
    public partial class Exel : Window
    {

        public Exel()
        {
            InitializeComponent();
        }


        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            MainWindow MW = new MainWindow();
            MW.Show();
            this.Close();
        }


        private void OpenRich_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog
            {
                Title = "Select an Excel file",
                Filters = { new CommonFileDialogFilter("Excel Files", "*.xlsx;*.xls") },
                EnsureFileExists = true
            };

            if (openFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string filePath = openFileDialog.FileName;

                Workbook VB = new Workbook();
                VB.LoadFromFile(filePath);

                Worksheet sheet = VB.Worksheets[0];
                CellRange locatedRange = sheet.AllocatedRange;

                var dataTable = sheet.ExportDataTable(locatedRange, true);
                dataGrid.ItemsSource = dataTable.DefaultView;

            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string columnName = textBox.Text;
            if (!string.IsNullOrWhiteSpace(columnName))
            {
                DataGridTextColumn textColumn = new DataGridTextColumn();
                textColumn.Header = columnName;
                textColumn.Binding = new System.Windows.Data.Binding(columnName);
                dataGrid.Columns.Add(textColumn);

                foreach (var item in dataGrid.Items)
                {
                    var dataItem = item as System.Dynamic.ExpandoObject;
                    if (dataItem != null)
                    {
                        ((IDictionary<string, object>)dataItem)[columnName] = "";
                    }
                }
            }
        }

        private void SaveRich_Click_1(object sender, RoutedEventArgs e)
        {
            var dataTable = dataGrid.ItemsSource as DataView;

            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = false;
            dialog.Title = "Выберите место сохранения файла Excel";
            dialog.DefaultExtension = ".xlsx";
            dialog.Filters.Add(new CommonFileDialogFilter("Файлы Excel", "*.xlsx"));

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok) 
            {
                Workbook wb = new Workbook();
                wb.Worksheets.Clear();
                Worksheet sheet = wb.Worksheets.Add("Лист 1");

                sheet.InsertDataView(dataTable, true, 1, 1);

                string filePath = dialog.FileName;
                wb.SaveToFile(filePath, Spire.Xls.FileFormat.Version2016);

                MessageBox.Show("Файл успешно сохранен по пути: " + filePath);
            }
        }

        private void SaveRich2_Click(object sender, RoutedEventArgs e)
        {
            pochta poch = new pochta();
            poch.Show();
        }
    }

}

