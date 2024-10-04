using Newtonsoft.Json;
using ConsoleApp1;
using ClosedXML.Excel;
using OfficeOpenXml;

class Program
{
    static string filePath = string.Empty;
    static string jsonString = string.Empty;
    static string jstring = string.Empty;
    static Root jsonObject = null;
    static string projectName = string.Empty;
    static string folderOrIIS = string.Empty;
    static Dictionary<string, int> pieChartData = null;
    static XLWorkbook workbook = new XLWorkbook();
    static void Main(string[] args)
    {
        try
        {
            Console.WriteLine("Please enter your Project path. Keep format of folders like- (ProjectName/ApplicationName(s) or IIS Server(s)/data/json)");
            filePath = Console.ReadLine();

            Console.WriteLine("Do you want json to Excel Conversion for IIS? Type Y for IIS and N for others");
            folderOrIIS = Console.ReadLine();

            string[] parts = filePath.Split('\\');

            projectName = parts[parts.Length - 1];
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }

        if (Directory.Exists(filePath))
        {
            // Recursively search for the results.json file within subdirectories
            // var pieChartData = SearchForResultsJson(filePath);
            var chartData = SearchForResultsJson(filePath);
            // Create Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //ExcelPackage excelPackage = new ExcelPackage();

            //Common.CreateExcelForCategory(pieChartData, filePath, excelPackage);           
            //excelPackage.SaveAs(new FileInfo(filePath + "\\" + "ApplicationCategoriesCount.xls"));

            XLWorkbook workbook = new XLWorkbook();
          
            //Common.CreateExcelForBarCategory(workbook, chartData.barChartData, filePath);
            Common.CreateExcelForPieCategory(workbook, chartData.pieChartData, chartData.barChartData, filePath);
            Console.WriteLine("ApplicationCategoriesCount excel is created.");

        }

        else
        {
            Console.WriteLine("Directory does not exist.");
        }


        static ChartData SearchForResultsJson(string folderPath)
        {
            try
            {
                // Get the subdirectories (folders) within the current directory
                string[] subDirectories = Directory.GetDirectories(folderPath);

                var chartData = new ChartData
                {
                    pieChartData = new List<Dictionary<string, int>>(),
                    barChartData = new List<Dictionary<string, int>>()
                };
                // var pieChartDatalist = new List<Dictionary<string, int>>();

                foreach (string subDirectory in subDirectories)
                {
                    // Check if the subdirectory contains a "data" folder
                    string dataFolderPath = Path.Combine(subDirectory, "data");

                    if (Directory.Exists(dataFolderPath))
                    {
                        // Check if the "data" folder contains a "json" file                       
                        string[] jsonFiles = Directory.GetFiles(dataFolderPath, "*.json");
                        string resultsJsonFilePath = Path.Combine(dataFolderPath, jsonFiles.First());


                        if (File.Exists(resultsJsonFilePath))
                        {
                            filePath = filePath.Trim().Trim('"');

                            jsonString = File.ReadAllText(resultsJsonFilePath);
                            if (!jsonString.Contains("results = \r\n{"))
                            {
                                jsonString = "results = " + jsonString;
                            }

                            jstring = jsonString.Replace("results = \r\n{", "{results : \r\n{");
                            jstring += "}";
                            jsonObject = JsonConvert.DeserializeObject<Root>(jstring);

                            if (jsonObject != null)
                            {
                                if (folderOrIIS == "Y")
                                {
                                    JsonToExcelL.CreateWorkbookForIIS(filePath, jsonObject, workbook, projectName, subDirectory, folderOrIIS);
                                }
                                else
                                {
                                    JsonToExcelL.CreateWorkbook(filePath, jsonObject, workbook, projectName, subDirectory);
                                }


                                var piechartdata = Common.FetchCategoriespieCount(jsonObject, workbook, filePath, subDirectory);
                                //pieChartDatalist.Add(piechartdata);
                                chartData.pieChartData.Add(piechartdata);
                                var barchartdata = Common.FetchCategoriesbarCount(jsonObject, workbook, filePath, subDirectory);

                                // Add bar chart data
                                chartData.barChartData.Add(barchartdata);
                            }
                        }

                    }

                    // Recursively search for results.json file in subdirectories
                    SearchForResultsJson(subDirectory);

                }
                return chartData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null;
            }
        }

    }

    class ChartData
    {
        public List<Dictionary<string, int>> pieChartData { get; set; }
        public List<Dictionary<string, int>> barChartData { get; set; }
    }
}

