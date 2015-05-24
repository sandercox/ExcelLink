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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using ExcelLink;

namespace ExcelLink.TestApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MyData data { get; set; }
        public Workbook workbook {get; set; }

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;

            data = new MyData();
            data.Firstname = "Firstname";
            data.Lastname = "Lastname";

            workbook = new Workbook("c:\\Users\\sande_000\\Documents\\testbook.xlsx", true);

            //Binding b = new Binding("Value2");
            //b.Source = xlRange;
            //b.Mode = BindingMode.TwoWay;
            //excelFirstname.SetBinding(TextBox.TextProperty, b);
            //Binding b = new Binding("Sheets[Sheet1].Cells[1,3].Value");
            //b.Source = workbook;
            //b.Mode = BindingMode.TwoWay;
            //excelFirstname.SetBinding(TextBox.TextProperty, b);
            //b = new Binding("Sheets[Sheet1].Cells[2,3].Value");
            //b.Source = workbook;
            //b.Mode = BindingMode.TwoWay;
            //excelLastname.SetBinding(TextBox.TextProperty, b);

            ExcelBinding.Bind(workbook.Sheets["Sheet1"].Cells[1, 2], data, "Firstname");
            ExcelBinding.Bind(workbook.Sheets["Sheet1"].Cells[2, 2], data, "Lastname");
        }
    }
}
