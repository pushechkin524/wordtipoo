using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Doc;
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
using System.Windows.Shapes;

namespace wordtipo
{
    /// <summary>
    /// Логика взаимодействия для CreateWord.xaml
    /// </summary>
    public partial class CreateWord : Window
    {
        public CreateWord()
        {
            InitializeComponent();
        }


        private void OpenRich_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("Word Documents", "*.docx"));
            dialog.Filters.Add(new CommonFileDialogFilter("All Files", "*.*"));

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string filePath = dialog.FileName;
                Document doc = new Document();
                doc.LoadFromFile(filePath);

                string rtfFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), "Конвертировали.rtf");
                doc.SaveToFile(rtfFilePath, FileFormat.Rtf);

                var range = new TextRange(myRichTextBox.Document.ContentStart, myRichTextBox.Document.ContentEnd);
                using (var fs = new FileStream(rtfFilePath, FileMode.OpenOrCreate))
                {
                    range.Load(fs, DataFormats.Rtf);
                }
            }
        }

        private void SaveRich_Click_1(object sender, RoutedEventArgs e)
        {

            // Open a dialog to select the location to save the DOCX file
            CommonSaveFileDialog dialog = new CommonSaveFileDialog
            {
                DefaultFileName = "Отформатированный файл в RTB",
                DefaultExtension = ".docx",
                Filters = { new CommonFileDialogFilter("Word Document", ".docx") }
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string docxFilePath = dialog.FileName;

                // Create a temporary RTF file
                string tempRtfFilePath = System.IO.Path.GetTempFileName();
                File.Move(tempRtfFilePath, tempRtfFilePath + ".rtf");
                tempRtfFilePath += ".rtf";

                // Save the content of RichTextBox to the temporary RTF file
                var range = new TextRange(myRichTextBox.Document.ContentStart, myRichTextBox.Document.ContentEnd);
                using (var fs = new FileStream(tempRtfFilePath, FileMode.Create))
                {
                    range.Save(fs, DataFormats.Rtf);
                }

                // Load the temporary RTF file and save it as DOCX
                Document doc = new Document();
                doc.LoadFromFile(tempRtfFilePath);
                doc.SaveToFile(docxFilePath, FileFormat.Docx);

                // Delete the temporary RTF file
                File.Delete(tempRtfFilePath);

                MessageBox.Show("File saved successfully!");
            }
        }



        private void SaveRich2_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            MainWindow MW = new MainWindow();
            MW.Show();
            this.Close();
        }
    }
}
