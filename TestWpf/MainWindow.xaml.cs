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

namespace TestWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            string rawcontent = @"<Grid xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width='*' />
        <ColumnDefinition Width='*' />
    </Grid.ColumnDefinitions>
    <TextBlock Text='id' Grid.Column='0'/>
    <Rectangle Fill='Black' Grid.Column='1' />
</Grid>";

            XmlReader xr = XmlReader.Create(new System.IO.StringReader(rawcontent));
            var control = XamlReader.Load(xr) as Grid;
            //this.Children.Add(control);

            this.AddChild(control);
        }
    }
}
