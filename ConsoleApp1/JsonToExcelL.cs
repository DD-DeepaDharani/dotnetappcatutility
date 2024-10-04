using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    public class Location
    {
        public string kind { get; set; }
        public string path { get; set; }
        public string snippet { get; set; }
    }
    public class RuleInstance
    {
        public string incidentId { get; set; }
        public string ruleId { get; set; }
        public string projectPath { get; set; }
        public string state { get; set; }
        public Location location { get; set; }
    }
    public class Links
    {
        public string title { get; set; }
        public string url { get; set; }
    }
    public class Rule
    {
        public string Id { get; set; }
        public string Description { get; set; }
        public string Label { get; set; }
        public string Severity { get; set; }
        public int Effort { get; set; }
        public List<Link> Links { get; set; }
    }
    public class Link
    {
        public string Title { get; set; }
        public string Url { get; set; }
    }
    public class Project
    {
        public string path { get; set; }
        public bool startingProject { get; set; }
        public int issues { get; set; }
        public int storyPoints { get; set; }
        public List<RuleInstance> ruleInstances { get; set; }
    }
    public class Results
    {
        public Stats stats { get; set; }
        public List<Project> projects { get; set; }
        public Dictionary<string, Rule> Rules { get; set; }

    }
    public class Stats
    {
        public Charts charts { get; set; }
    }
    public class Charts
    {
        public Dictionary<string, int> severity { get; set; }
        public Dictionary<string, int> category { get; set; }
    }

    public class Root
    {
        public Results results { get; set; }

    }
    public class JsonToExcelL
    {
        static string substringFileName = string.Empty;
        static string propath = string.Empty;
        public static void CreateWorkbook(string filePath, Root jsonObject, XLWorkbook workbook, string projectName, string subDirectory)
        {
            try
            {
                substringFileName = subDirectory.Split('\\').LastOrDefault();

                var worksheet = workbook.Worksheets.Add(substringFileName);
                Common.AddingHeadings(worksheet);

                int row = 2;
                var result = string.Empty;
                int index = 0;
                string pathValue = string.Empty;
                string path = string.Empty;
                string trimPath = string.Empty;

                // Add data from Runtime
                if (jsonObject != null && jsonObject.results.Rules != null)
                {
                    foreach (var kvp in jsonObject.results.Rules)
                    {
                        var data = kvp.Value;
                        index = data.Description.IndexOf("\n\n");
                        result = index != -1 ? data.Description.Substring(index) : data.Description;
                        int lastIndex = data.Description.IndexOf("\n\n");
                        string resultLBl;
                        resultLBl = lastIndex != -1 ? data.Description.Substring(0, lastIndex) : data.Description;

                        Common.SetPriorityAccordingToSeverity(data, worksheet, row, resultLBl, result);

                        if (jsonObject != null && jsonObject.results.projects != null)
                        {
                            for (int k = 0; k < jsonObject.results.projects.Count; k++)
                            {
                                if (jsonObject.results.projects[k].ruleInstances.Count > 0)
                                {

                                    var projectsWithMatchingRule = jsonObject.results.projects[k].ruleInstances
            .Where(project => project.ruleId == data.Id).ToList();

                                    if (projectsWithMatchingRule.Count > 0)
                                    {
                                        int countVal = 0;
                                        if (projectsWithMatchingRule.Count >= 40)
                                        {
                                            countVal = 40;
                                        }

                                        else
                                        {
                                            countVal = projectsWithMatchingRule.Count;
                                        }
                                        // Assuming you want to loop through all ruleInstances for each project
                                        for (int j = 0; j < countVal; j++)

                                        {

                                            var lastIndexs = projectsWithMatchingRule[j].location.path.LastIndexOf("\\");
                                            string resultLBl1;
                                            resultLBl1 = lastIndexs != -1 ? projectsWithMatchingRule[j].location.path.Substring(lastIndexs + 1) : projectsWithMatchingRule[j].location.path;

                                            if (!string.IsNullOrEmpty(resultLBl1) && resultLBl1 != pathValue.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).LastOrDefault())
                                            {

                                                if (!pathValue.Contains(resultLBl1))
                                                {
                                                    pathValue += resultLBl1 + Environment.NewLine;
                                                }


                                            }
                                            if (projectsWithMatchingRule[j].location.snippet != null)
                                            {
                                                if (projectsWithMatchingRule[j].location.snippet.ToLower().Contains("data source") || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("password")
                                           || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("value =") || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("value="))
                                                {
                                                    // Mask    
                                                    string maskedDataSourceConnectionString = Regex.Replace(projectsWithMatchingRule[j].location.snippet.ToLower(), @"(?<=data source=)([^;]+)", "*****");
                                                    maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, "(?<=password=)([^;]+)", "****");
                                                    maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=user id=)([^;]+)", "*****");
                                                    maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=initial catalog=)([^;]+)", "*****");
                                                    //maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=metadata=)([^;]+)", "*****");
                                                    maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=value=)([^;]+)", "*****");

                                                    int maxLength = 32767;
                                                    if (maskedDataSourceConnectionString.Length > maxLength)
                                                    {
                                                        worksheet.Cell(row, 5).Value = maskedDataSourceConnectionString.Substring(0, maxLength);
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cell(row, 5).Value = maskedDataSourceConnectionString;
                                                    }
                                                }




                                                else
                                                {
                                                    int maxLength = 32767;
                                                    if (projectsWithMatchingRule[0].location.snippet.Length > maxLength)
                                                    {
                                                        worksheet.Cell(row, 5).Value = projectsWithMatchingRule[0].location.snippet.Substring(0, maxLength);
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cell(row, 5).Value = projectsWithMatchingRule[0].location.snippet;
                                                    }

                                                }
                                            }
                                        }

                                        worksheet.Cell(row, 4).Value = pathValue.TrimEnd('\r', '\n');

                                        path += jsonObject.results.projects[k].path;
                                        worksheet.Cell(row, 8).Value += path.Substring(path.LastIndexOf("\\") + 1) + "\n";
                                    }
                                }
                            }
                            row++;
                            pathValue = string.Empty;

                        }
                    }

                    Common.SetAllignmentsAndSave(worksheet, workbook, filePath, projectName, substringFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception exists in File " + substringFileName + " " + ex.Message);
            }
        }

        public static void CreateWorkbookForIIS(string filePath, Root jsonObject, XLWorkbook workbook, string projectName, string subDirectory, string folderOrIIS)
        {
            try
            {
                substringFileName = subDirectory.Split('\\').LastOrDefault();

                var projects = jsonObject.results.projects.Where(project => project.ruleInstances.Count > 0);
                {
                    foreach (var project in projects)
                    {
                        string projectpath = project.path;
                        propath = projectpath.Replace('/', '.');

                        var worksheet = workbook.Worksheets.Add(propath);
                        Common.AddingHeadings(worksheet);

                        int row = 2;
                        var result = string.Empty;
                        int index = 0;
                        string pathValue = string.Empty;
                        string path = string.Empty;
                        string trimPath = string.Empty;

                        // Add data from Runtime
                        if (jsonObject != null && jsonObject.results.Rules != null)
                        {
                            foreach (var kvp in jsonObject.results.Rules)
                            {
                                var data = kvp.Value;
                                index = data.Description.IndexOf("\n\n");
                                result = index != -1 ? data.Description.Substring(index) : data.Description;
                                int lastIndex = data.Description.IndexOf("\n\n");
                                string resultLBl;
                                resultLBl = lastIndex != -1 ? data.Description.Substring(0, lastIndex) : data.Description;

                                Common.SetPriorityAccordingToSeverity(data, worksheet, row, resultLBl, result);


                                if (jsonObject != null && jsonObject.results.projects != null)
                                {
                                    for (int k = 0; k < jsonObject.results.projects.Count; k++)
                                    {
                                        if (jsonObject.results.projects[k].ruleInstances.Count > 0)
                                        {

                                            var projectsWithMatchingRule = jsonObject.results.projects[k].ruleInstances
                    .Where(project => project.ruleId == data.Id).ToList();

                                            if (projectsWithMatchingRule.Count > 0)
                                            {
                                                int countVal = 0;
                                                if (projectsWithMatchingRule.Count >= 40)
                                                {
                                                    countVal = 40;
                                                }

                                                else
                                                {
                                                    countVal = projectsWithMatchingRule.Count;
                                                }
                                                // Assuming you want to loop through all ruleInstances for each project
                                                for (int j = 0; j < countVal; j++)

                                                {

                                                    var lastIndexs = projectsWithMatchingRule[j].location.path.LastIndexOf("\\");
                                                    string resultLBl1;
                                                    resultLBl1 = lastIndexs != -1 ? projectsWithMatchingRule[j].location.path.Substring(lastIndexs + 1) : projectsWithMatchingRule[j].location.path;

                                                    if (!string.IsNullOrEmpty(resultLBl1) && resultLBl1 != pathValue.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).LastOrDefault())
                                                    {

                                                        if (!pathValue.Contains(resultLBl1))
                                                        {
                                                            pathValue += resultLBl1 + Environment.NewLine;
                                                        }

                                                    }

                                                    if (projectsWithMatchingRule[j].location.snippet.ToLower().Contains("data source") || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("password")
                                                   || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("value =") || projectsWithMatchingRule[j].location.snippet.ToLower().Contains("value="))
                                                    {
                                                        // Mask    
                                                        string maskedDataSourceConnectionString = Regex.Replace(projectsWithMatchingRule[j].location.snippet.ToLower(), @"(?<=data source=)([^;]+)", "*****");
                                                        maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, "(?<=password=)([^;]+)", "****");
                                                        maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=user id=)([^;]+)", "*****");
                                                        maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=initial catalog=)([^;]+)", "*****");
                                                        //maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=metadata=)([^;]+)", "*****");
                                                        maskedDataSourceConnectionString = Regex.Replace(maskedDataSourceConnectionString, @"(?<=value=)([^;]+)", "*****");

                                                        int maxLength = 32767;
                                                        if (maskedDataSourceConnectionString.Length > maxLength)
                                                        {
                                                            worksheet.Cell(row, 5).Value = maskedDataSourceConnectionString.Substring(0, maxLength);
                                                        }
                                                        else
                                                        {
                                                            worksheet.Cell(row, 5).Value = maskedDataSourceConnectionString;
                                                        }
                                                    }


                                                    else
                                                    {
                                                        int maxLength = 32767;
                                                        if (projectsWithMatchingRule[0].location.snippet.Length > maxLength)
                                                        {
                                                            worksheet.Cell(row, 5).Value = projectsWithMatchingRule[0].location.snippet.Substring(0, maxLength);
                                                        }
                                                        else
                                                        {
                                                            worksheet.Cell(row, 5).Value = projectsWithMatchingRule[0].location.snippet;
                                                        }


                                                    }
                                                }

                                                worksheet.Cell(row, 4).Value = pathValue.TrimEnd('\r', '\n');
                                                worksheet.Cell(row, 8).Value = propath;

                                            }
                                        }
                                    }
                                    row++;
                                    pathValue = string.Empty;

                                }
                            }

                            Common.SetAllignmentsAndSave(worksheet, workbook, filePath, projectName, propath);
                        }
                    }

                }



            }
            catch (Exception ex)
            {

                Console.WriteLine("Exception exists in File " + propath + ex.Message);
            }
        }
    }
}
