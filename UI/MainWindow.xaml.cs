using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace UI
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string targetFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"20180605検査員の作業範囲.xlsx");
            fileLocation.Text = targetFile;
        }

        private void readFile_Click(object sender, RoutedEventArgs e)
        {
            string str = ExcelToWpf.Program.generateTestGridString(fileLocation.Text);
            parsedExcelContentViewer.Text = str;
        }

        private void generateWindow_Click(object sender, RoutedEventArgs e)
        {
            if(String.IsNullOrWhiteSpace(parsedExcelContentViewer.Text))
            {
                return;
            }

            EmptyWindow emp = new EmptyWindow(parsedExcelContentViewer.Text);
            emp.Owner = this;
            emp.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            emp.ShowDialog();
        }
    }
}
