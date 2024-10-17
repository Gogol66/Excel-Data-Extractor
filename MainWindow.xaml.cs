using LiveCharts;
using LiveCharts.Configurations;
using LiveCharts.Wpf;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
namespace TestChartWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public SeriesCollection Values { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            LoadData();
            DataContext = this;
        }
        private void LoadData()
        {
            var data = ReadExcelData("C:/Users/saura/OneDrive/Desktop/file_example_XLSX_5000.xlsx");

            var mapper = Mappers.Xy<KeyValuePair<double, double>>().X(model => model.Key).Y(model => model.Value); // Y value

            Charting.For<KeyValuePair<string, string>>(mapper);

            Values = new SeriesCollection();

            // Add data points to the chart
            foreach (var point in data)
            {
                Values.Add(new LineSeries
                {
                    Values = new ChartValues<string> { point.Value },
                    Title = point.Key // You can customize the title or use other properties
                });
            }
        }

        public List<KeyValuePair<string, string>> ReadExcelData(string filePath)
        {
            var data = new List<KeyValuePair<string, string>>();

            // Ensure that EPPlus can read Excel files by setting the ExcelPackage.LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
                int rowCount = worksheet.Dimension.Rows;

                // Assuming the first row is a header, start reading from the second row
                for (int row = 2; row <= rowCount; row++)
                {
                    var xValue = worksheet.Cells[row, 6].Text; // First column (string for X-axis)
                    var yValue = worksheet.Cells[row, 8].Text; // Second column (double for Y-axis)

                    // Add the data point to the list
                    data.Add(new KeyValuePair<string, string>(xValue, yValue));
                }
            }

            return data; // Return the list of data points
        }

    }
}
