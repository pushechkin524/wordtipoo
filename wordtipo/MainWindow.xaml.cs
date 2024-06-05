using Microsoft.Win32;
using System;
using Spire.Doc;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
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

namespace wordtipo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateWord_Click(object sender, object e)
        {
            CreateWord CW = new CreateWord();
            CW.Show();
            this.Close();
        }

        private void CreateExel_Click(object sender, object e)
        {
            Exel EX = new Exel();
            EX.Show();
            this.Close();
        }
    }
}
