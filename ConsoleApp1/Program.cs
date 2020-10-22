using OfficeOpenXml;
using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ConsoleApp1
{
    class Program
    {
        static int column = 1;
        static int row = 1;

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            DoSomething();
        }

        static void DoSomething()
        {
            var src = File.ReadAllText("templateAll.json");
            using var document = System.Text.Json.JsonDocument.Parse(src);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excel = new FileInfo("output.xlsx");
            if (excel.Exists) excel.Delete();

            using ExcelPackage package = new ExcelPackage(excel);
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("resource");
            var resources = document.RootElement.GetProperty("resources");

            foreach (var resource in resources.EnumerateArray())
            {
                WriteResourceToSheet(sheet, resource);
            }

            package.Save();
        }

        static void WriteResourceToSheet(ExcelWorksheet sheet, JsonElement jsonElement)
        {
            if (jsonElement.ValueKind == JsonValueKind.Object)
            {
                foreach (var property in jsonElement.EnumerateObject())
                {
                    sheet.Cells[row, column].Value = property.Name;
                    
                    switch (property.Value.ValueKind)
                    {
                        case JsonValueKind.Object:
                            row++;
                            foreach (var childProperty in property.Value.EnumerateObject())
                            {
                                switch (childProperty.Value.ValueKind)
                                {
                                    case JsonValueKind.Array:
                                        foreach (var childItem in childProperty.Value.EnumerateArray())
                                        {
                                            WriteResourceToSheet(sheet, childItem);
                                        }
                                        break;
                                    case JsonValueKind.Object:
                                        foreach (var obj in childProperty.Value.EnumerateObject())
                                        {
                                            WriteResourceToSheet(sheet, obj);
                                        }
                                        break;
                                    default:
                                        WriteResourceToSheet(sheet, childProperty);
                                        break;
                                }
                            }
                            break;
                        case JsonValueKind.Array:
                            foreach (var childItem in property.Value.EnumerateArray())
                            {
                                WriteResourceToSheet(sheet, childItem);
                            }
                            break;
                        default:
                            sheet.Cells[row, column + 1].Value = property.Value;
                            row++;
                            break;
                    }                 
                }
            }
            else
            {
                sheet.Cells[row, column + 1].Value = jsonElement.ToString();
                row++;
            }
        }

        static void WriteResourceToSheet(ExcelWorksheet sheet, JsonProperty jsonProperty)
        {
            sheet.Cells[row, column].Value = jsonProperty.Name;
            sheet.Cells[row, column + 1].Value = jsonProperty.Value;
            row++;
        }
    }
}
