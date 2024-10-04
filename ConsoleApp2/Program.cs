using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

class Program
{
    static void Main(string[] args)
    {
        // Example data
        var pieChartData = new List<Dictionary<string, double>>
        {
            new Dictionary<string, double>
            {
                {"Category 1", 30},
                {"Category 2", 20},
                {"Category 3", 50}
            },
            new Dictionary<string, double>
            {
                {"Category A", 40},
                {"Category B", 10},
                {"Category C", 50}
            }
            // Add more datasets here
        };
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            int rowStart = 1;
            int chartNumber = 1;

            foreach (var data in pieChartData)
            {
                int colStart = 1;
                worksheet.Cells[rowStart, colStart].Value = "Category";
                worksheet.Cells[rowStart, colStart + 1].Value = "Value";

                int row = rowStart + 1;
                foreach (var entry in data)
                {
                    worksheet.Cells[row, colStart].Value = entry.Key;
                    worksheet.Cells[row, colStart + 1].Value = entry.Value;
                    row++;
                }

                var pieChart = worksheet.Drawings.AddChart($"pieChart{chartNumber}", eChartType.Pie) as ExcelPieChart;
                pieChart.Title.Text = $"Pie Chart {chartNumber}";
                pieChart.Series.Add(worksheet.Cells[rowStart + 1, colStart + 1, row - 1, colStart + 1], worksheet.Cells[rowStart + 1, colStart, row - 1, colStart]);
                pieChart.SetPosition(rowStart - 1, 0, colStart + 2, 0);
                pieChart.SetSize(250, 250);

                rowStart = row + 10; // Move to the next row after the current data and some space for the chart
                chartNumber++;
            }

            // Save the Excel file
            FileInfo file = new FileInfo(@"C:\Users\v-sonalrenge\OneDrive - Microsoft\Reports\file.xls");
            package.SaveAs(file);
        }

    }
}

