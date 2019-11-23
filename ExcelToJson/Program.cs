using OfficeOpenXml;
using System;
using System.IO;

namespace JsonReplacer
{
    class Program
    {
        static void Main(string[] args) {
            Console.WriteLine("Reading File");
            JsonGenerate();
        }

        static void JsonGenerate() {
            var filePath = @"C:\temp\Values.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++) {
                    string json = File.ReadAllText(@"C:\temp\sample.json");
                    dynamic jsonObj = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                    var filename = "";

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++) {
                        string value = worksheet.Cells[row, col].Value.ToString();
                        switch (col) {
                            case 1:
                                jsonObj["glossary"]["title"] = value;
                                break;
                            case 2:
                                jsonObj["glossary"]["GlossDiv"]["title"] = value;
                                break;
                            case 3:
                                filename = value;
                                break;
                            case 4:
                                jsonObj["glossary"]["GlossDiv"]["GlossList"]["GlossEntry"]["ID"] = value;
                                break;
                            default:
                                break;
                        }
                    }
                    string output = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObj, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText($@"C:\temp\{filename}.json", output);
                }
            }

        }
    }
}
