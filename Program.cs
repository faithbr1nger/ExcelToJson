using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace JsonReplacer
{
    class Program
    {
        static void Main(string[] args) {
            Console.WriteLine("Reading File");

            Console.WriteLine(@"Enter file path (e.g. C:\NZME.Deployments\sites\parameters\):");
            string parametersFolderPath = Console.ReadLine();
            try
            {
                JsonGenerate(parametersFolderPath);
            }catch(Exception e)
            {
                Console.WriteLine($@"Error: {e.Message}. Press Any Key to quit...");
                Console.ReadLine();
            }
            Console.WriteLine("Press Any Key to continue...");
            Console.ReadLine();
        }

        static void JsonGenerate(string parametersFolderPath) {
            var filePath = @".\Input\Values.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++) {
                    string json = File.ReadAllText(@".\Input\sample.json");
                    dynamic jsonObj = Newtonsoft.Json.JsonConvert.DeserializeObject(json, typeof(object));
                    string frontOrBackoffice = "";
                    string environment = "";
                    string site = "";

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        string value = worksheet.Cells[row, col].Value.ToString();
                        switch (col) 
                        {
                            case 1:
                                jsonObj["parameters"]["actionGroupShortName"]["value"] = value;
                                break;
                            case 2:
                                frontOrBackoffice = value;
                                jsonObj["parameters"]["instanceName"]["value"] = value;
                                break;
                            case 3:
                                environment = value;
                                jsonObj["parameters"]["environment"]["value"] = value;
                                break;
                            case 4:
                                jsonObj["parameters"]["location"]["value"] = value;
                                break;
                            case 5:
                                jsonObj["parameters"]["pingTests"]["value"][0]["name"] = value;
                                jsonObj["parameters"]["pingTests"]["value"][0]["syntheticMonitorName"] = value;
                                jsonObj["parameters"]["pingTests"]["value"][0]["webTestName"] = value;
                                jsonObj["parameters"]["pingTests"]["value"][0]["metricAlert"]["alertName"] = value;
                                break;
                            case 6:
                                jsonObj["parameters"]["pingTests"]["value"][0]["webTestRequestUrl"] = value;
                                break;
                            case 7:
                                jsonObj["parameters"]["pingTests"]["value"][0]["webTestExpectedHttpStatusCode"] = value;
                                break;
                            case 8:
                                jsonObj["parameters"]["pingTests"]["value"][0]["webTestExpectedText"] = value;
                                break;
                            case 9:
                                jsonObj["parameters"]["pingTests"]["value"][0]["applicationInsightsResourceName"] = value;
                                break;
                            case 10:
                                jsonObj["parameters"]["pingTests"]["value"][0]["metricAlert"]["alertDescription"] = value;
                                break;
                            case 11:
                                jsonObj["parameters"]["pingTests"]["value"][0]["metricAlert"]["alertSeverity"] = value;
                                break;
                            case 12:
                                site = value;
                                jsonObj["parameters"]["contextName"]["value"] = value;
                                break;
                            case 13:
                                jsonObj["parameters"]["emailList"]["value"] = Regex.Replace(value, @"\t|\n|\r", "");
                                break;
                            default:
                                
                                break;
                        }
                    }
                    jsonObj["parameters"]["pingTests"]["value"][0]["webTestId"] = Guid.NewGuid().ToString();
                    jsonObj["parameters"]["pingTests"]["value"][0]["webTestRequestId"] = Guid.NewGuid().ToString();
                    string output = JsonConvert.SerializeObject(jsonObj, Formatting.Indented);
                    
                    output = Regex.Replace(output, @"\\", "");
                    output = Regex.Replace(output, @"\""\[", "[");
                    output = Regex.Replace(output, @"\]\""", "]");
                    output = JValue.Parse(output).ToString(Formatting.Indented);
                    Directory.CreateDirectory($@"{parametersFolderPath}\{environment}\{site}\{frontOrBackoffice}\");
                    File.WriteAllText($@"{parametersFolderPath}\{environment}\{site}\{frontOrBackoffice}\azure-monitor-deploy.parameters.json", output);
                    Console.WriteLine($@"Json successfully generated in {parametersFolderPath}\{environment}\{site}\{frontOrBackoffice}\azure-monitor-deploy.parameters.json");
                }
            }

        }
    }
}
