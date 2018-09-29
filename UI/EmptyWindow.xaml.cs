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
using System.Windows.Shapes;
using System.Xml;

namespace UI
{
    /// <summary>
    /// Interaction logic for EmptyWindow.xaml
    /// </summary>
    public partial class EmptyWindow : Window
    {
        public EmptyWindow(string text)
        {
            InitializeComponent();

            XmlReader xr = XmlReader.Create(new System.IO.StringReader(text));

            ParserContext context = new ParserContext();
            context.XmlnsDictionary.Add("", "http://schemas.microsoft.com/winfx/2006/xaml/presentation");
            context.XmlnsDictionary.Add("x", "http://schemas.microsoft.com/winfx/2006/xaml");


            //var control = XamlReader.Load(xr) as Grid;

            var control =  XamlReader.Parse(text, context) as Grid;

            this.AddChild(control);
        }
    }
}
