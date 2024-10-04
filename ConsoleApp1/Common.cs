using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Reflection.Metadata;


namespace ConsoleApp1
{
    public class Common
    {
        public static void SetPriorityAccordingToSeverity(Rule data, IXLWorksheet worksheet, int row, string resultLBl, string result)
        {
            worksheet.Cell(row, 1).Value = data.Label;   // Column 1, row 1
            worksheet.Cell(row, 2).Value = data.Severity; // Column 2, row 1
            worksheet.Cell(row, 6).Value = resultLBl;    // Column 6, row 1
            worksheet.Cell(row, 7).Value = result;       // Column 7, row 1

            switch (data.Severity)
            {
                case "Mandatory":
                    worksheet.Cell(row, 3).Value = "Pre check to be met"; // Column 2, row 1
                    worksheet.Cell(row, 9).Value = "P0"; // Column 8, row 1
                    break;

                case "Optional" or "Information":
                    worksheet.Cell(row, 3).Value = "Good to Have"; // Column 2, row 1
                    worksheet.Cell(row, 9).Value = "P2"; // Column 8, row 1

                    break;

                case "Potential":
                    worksheet.Cell(row, 3).Value = "Best Practice"; // Column 2, row 1
                    worksheet.Cell(row, 9).Value = "P1"; // Column 8, row 1
                    break;
                default:
                    break;
            }
        }

        public static void AddingHeadings(IXLWorksheet worksheet)
        {
            var headerRange = worksheet.Range("A1:L1"); // Assuming header row is from column A to E
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


            worksheet.Column(1).Width = 40; // Column A
            worksheet.Column(2).Width = 15;// Column B

            worksheet.Column(3).Width = 15;// Column C
            worksheet.Column(4).Width = 30;// Column D
            worksheet.Column(5).Width = 50; // Column E
            worksheet.Column(6).Width = 40; // Column F
            worksheet.Column(7).Width = 50; // Column G
            worksheet.Column(8).Width = 30;// Column H
            worksheet.Column(9).Width = 15;// Column I
            worksheet.Column(10).Width = 15;// Column J
            worksheet.Column(11).Width = 15;// Column k

            worksheet.Cell(1, 1).Value = "Description";
            worksheet.Cell(1, 2).Value = "Severity";
            worksheet.Cell(1, 3).Value = "Validation check";
            worksheet.Cell(1, 4).Value = "Reference Location";
            worksheet.Cell(1, 5).Value = "Source";
            worksheet.Cell(1, 6).Value = "Current State";
            worksheet.Cell(1, 7).Value = "Recommendation";
            worksheet.Cell(1, 8).Value = "Project";
            worksheet.Cell(1, 9).Value = "Priority";
            worksheet.Cell(1, 10).Value = "Customer Remediation Decision";
            worksheet.Cell(1, 11).Value = "ETA";
            worksheet.Cell(1, 12).Value = "Assigned To";

        }

        public static Dictionary<string, int> FetchCategoriespieCount(Root jsonObject, XLWorkbook workbook, string filePath, string subDirectory)
        {
            try
            {

                if (jsonObject.results.projects.Count() > 0)
                {

                    var piechartdata = new Dictionary<string, int>();

                    foreach (var item in jsonObject.results.stats.charts.severity)
                    {
                        // if (item.Value >0)
                        piechartdata[item.Key] = item.Value;
                    }


                    return piechartdata;
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }
        }

        public static Dictionary<string, int> FetchCategoriesbarCount(Root jsonObject, XLWorkbook workbook, string filePath, string subDirectory)
        {
            try
            {

                if (jsonObject.results.projects.Count() > 0)
                {

                    var barchartdata = new Dictionary<string, int>();

                    foreach (var item in jsonObject.results.stats.charts.category)
                    {
                        // if (item.Value >0)
                        barchartdata[item.Key] = item.Value;
                    }


                    return barchartdata;
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }
        }

        public static string GetFileName(string filePath)
        {
            var result = filePath.Substring(filePath.LastIndexOf('\\') + 1);
            string excelname = result + "_net_apcat_asmt_analysis.xlsx";
            string Path = System.IO.Path.Combine(filePath, excelname);
            return Path;
        }

        public static void CreateExcelForPieCategory(XLWorkbook workbook1, List<Dictionary<string, int>> pieChartData, List<Dictionary<string, int>> barChartData, string filePath1)
        {

            string filePath = GetFileName(filePath1);

            using (var workbook = new XLWorkbook(filePath))
            {
                // Save the workbook to a memory stream
                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.Position = 0;

                    // Open the workbook with EPPlus
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        // Add a worksheet
                        var worksheet1 = package.Workbook.Worksheets.Add("Dashboard");
                        int rowStart = 3;
                        int rowStartBar = 3;
                        int i = 0;
                        int j = 0;
                        int b = 0;

                        string[] subdirectories = Directory.GetDirectories(filePath1);
                        foreach (string subdirectory in subdirectories)
                        {
                            subdirectories[j] = System.IO.Path.GetFileName(subdirectory);
                            j++;
                        }

                        foreach (var data in pieChartData)
                        {
                            int colStart = 1;


                            //worksheet1.Cells[rowStart, colStart].Value = "Category";
                            //worksheet1.Cells[rowStart, colStart + 1].Value = "Value";

                            int row = rowStart + 1;
                            foreach (var kvp in data)
                            {
                                worksheet1.Cells[row, colStart].Value = kvp.Key;
                                worksheet1.Cells[row, colStart + 1].Value = kvp.Value;
                                worksheet1.Cells[row, colStart].Style.Hidden = true;
                                worksheet1.Cells[row, colStart + 1].Style.Hidden = true;
                                worksheet1.Cells[row, colStart].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                worksheet1.Cells[row, colStart + 1].Style.Font.Color.SetColor(System.Drawing.Color.White);

                                row++;                           
                            }
                            worksheet1.Cells[1, 4].Value = "Summary for Severity";
                            worksheet1.Cells[1, 12].Value = "Summary for Categories";
                            worksheet1.Cells[1,4].Style.Font.Bold = true;
                            worksheet1.Cells[1, 12].Style.Font.Bold = true;

                            var pieChart = worksheet1.Drawings.AddChart($"{subdirectories[i]}", eChartType.Pie) as ExcelPieChart;

                            pieChart.Title.Text = subdirectories[i];
                            pieChart.Series.Add(worksheet1.Cells[rowStart + 1, colStart + 1, row - 1, colStart + 1], worksheet1.Cells[rowStart + 1, colStart, row - 1, colStart]);
                            pieChart.SetPosition(rowStart - 1, 0, colStart + 1, 0);
                            pieChart.SetSize(250, 200);
                            pieChart.StyleManager.SetChartStyle(ePresetChartStyle.PieChartStyle1, ePresetChartColors.ColorfulPalette4);
                            pieChart.DataLabel.ShowValue = true; // Display the values as data labels on the chart                         

                            rowStart = row + 10; // Move to the next row after the current data and some space for the chart
                            i++;
                        }

                        foreach (var data in barChartData)
                        {
                            int colStartBar = 7;

                            //worksheet1.Cells[rowStartBar, colStartBar].Value = "Category";
                            //worksheet1.Cells[rowStartBar, colStartBar + 1].Value = "Value";

                            int rowBar = rowStartBar + 1;
                            foreach (var kvp in data)
                            {
                                worksheet1.Cells[rowBar, colStartBar].Value = kvp.Key;
                                worksheet1.Cells[rowBar, colStartBar + 1].Value = kvp.Value;

                                worksheet1.Cells[rowBar, colStartBar].Style.Hidden = true;
                                worksheet1.Cells[rowBar, colStartBar + 1].Style.Hidden = true;

                                worksheet1.Cells[rowBar, colStartBar].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                worksheet1.Cells[rowBar, colStartBar + 1].Style.Font.Color.SetColor(System.Drawing.Color.White);

                                rowBar++;
                            }

                            var pieChart = worksheet1.Drawings.AddChart($"{subdirectories[b]}" + "barChart", eChartType.ColumnClustered) as ExcelBarChart;
                            pieChart.Title.Text = subdirectories[b];
                            pieChart.Series.Add(worksheet1.Cells[rowStartBar + 1, colStartBar + 1, rowBar - 1, colStartBar + 1], worksheet1.Cells[rowStartBar + 1, colStartBar, rowBar - 1, colStartBar]);

                            for (int s = 0; s < pieChart.Series.Count(); s++)
                            {
                                var bar = pieChart.Series[s] as ExcelBarChartSerie;
                                bar.Fill.Color = System.Drawing.Color.FromArgb(79, 129, 189); // Example color (RGB values)
                            }

                            pieChart.XAxis.MajorGridlines.Fill.Color = System.Drawing.Color.Transparent;
                            pieChart.XAxis.MajorTickMark = eAxisTickMark.None;
                            pieChart.XAxis.MinorTickMark = eAxisTickMark.None;
                            pieChart.XAxis.Border.Fill.Color = System.Drawing.Color.LightGray;
                            pieChart.YAxis.MajorGridlines.Fill.Color = System.Drawing.Color.LightGray;
                            pieChart.YAxis.MajorTickMark = eAxisTickMark.None;
                            pieChart.YAxis.MinorTickMark = eAxisTickMark.None;
                            pieChart.YAxis.Border.Fill.Color = System.Drawing.Color.Transparent;

                            pieChart.SetPosition(rowStartBar - 1, 0, colStartBar + 1, 0);
                            pieChart.SetSize(450, 200);
                            pieChart.DataLabel.ShowValue = true;

                            rowStartBar = rowBar + 10; // Move to the next row after the current data and some space for the chart
                            b++;
                        }

                        // Save the modified workbook
                        package.SaveAs(new FileInfo(filePath));
                    }
                }
            }

            Console.WriteLine("Excel file with pie chart created successfully!");
        }

        public static void CreateExcelForBarCategory(XLWorkbook workbook1, List<Dictionary<string, int>> barChartData, string filePath1)
        {
            string filePath = GetFileName(filePath1);

            using (var workbook = new XLWorkbook(filePath))
            {
                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.Position = 0;

                    // Open the workbook with EPPlus
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        // Add a worksheet

                        var worksheet1 = package.Workbook.Worksheets.Add("Summary for categories");
                        int rowStart = 1;
                        int i = 0;
                        int j = 0;

                        string[] subdirectories = Directory.GetDirectories(filePath1);
                        foreach (string subdirectory in subdirectories)
                        {
                            subdirectories[j] = System.IO.Path.GetFileName(subdirectory);
                            j++;
                        }

                        foreach (var data in barChartData)
                        {
                            int colStart = 10;


                            worksheet1.Cells[rowStart, colStart].Value = "Category";
                            worksheet1.Cells[rowStart, colStart + 1].Value = "Value";

                            int row = rowStart + 1;
                            foreach (var kvp in data)
                            {
                                worksheet1.Cells[row, colStart].Value = kvp.Key;
                                worksheet1.Cells[row, colStart + 1].Value = kvp.Value;
                                row++;
                            }

                            var pieChart = worksheet1.Drawings.AddChart($"{subdirectories[i]}", eChartType.ColumnClustered) as ExcelBarChart;
                            pieChart.Title.Text = subdirectories[i];
                            pieChart.Series.Add(worksheet1.Cells[rowStart + 1, colStart + 1, row - 1, colStart + 1], worksheet1.Cells[rowStart + 1, colStart, row - 1, colStart]);

                            for (int s = 0; s < pieChart.Series.Count(); s++)
                            {
                                var bar = pieChart.Series[s] as ExcelBarChartSerie;
                                bar.Fill.Color = System.Drawing.Color.FromArgb(79, 129, 189); // Example color (RGB values)
                            }

                            pieChart.XAxis.MajorGridlines.Fill.Color = System.Drawing.Color.Transparent;
                            pieChart.XAxis.MajorTickMark = eAxisTickMark.None;
                            pieChart.XAxis.MinorTickMark = eAxisTickMark.None;
                            pieChart.XAxis.Border.Fill.Color = System.Drawing.Color.LightGray;
                            pieChart.YAxis.MajorGridlines.Fill.Color = System.Drawing.Color.LightGray;
                            pieChart.YAxis.MajorTickMark = eAxisTickMark.None;
                            pieChart.YAxis.MinorTickMark = eAxisTickMark.None;
                            pieChart.YAxis.Border.Fill.Color = System.Drawing.Color.Transparent;

                            pieChart.SetPosition(rowStart - 1, 0, colStart + 2, 0);
                            pieChart.SetSize(600, 250);

                            rowStart = row + 10; // Move to the next row after the current data and some space for the chart
                            i++;
                        }
                        package.SaveAs(new FileInfo(filePath));
                    }
                }
            }
        }

        public static void SetAllignmentsAndSave(IXLWorksheet worksheet, XLWorkbook workbook, string filePath, string projectName, string propath)
        {
            worksheet.CellsUsed().Style.Alignment.WrapText = true;
            worksheet.Row(1).Height = 15;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Rows().Style.Font.FontName = "Calibri"; // Set font name
            worksheet.Rows().Style.Font.FontSize = 11;

            for (int m = 2; m <= worksheet.Rows().Count(); m++)
            {
                worksheet.Row(m).Height = 80;
            }

            var usedRange = worksheet.CellsUsed();
            usedRange.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            usedRange.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            usedRange.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
            usedRange.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            workbook.SaveAs(filePath + "\\" + projectName + "_net_apcat_asmt_analysis.xlsx");

            Console.WriteLine(propath + " file created successfully.");

        }
    }
}
