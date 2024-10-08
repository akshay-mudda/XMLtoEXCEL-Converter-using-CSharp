using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace XMLtoEXCEL_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Read source and destination paths from the app.config file
                string sourceDirectory = ConfigurationManager.AppSettings["SourceDirectory"];
                string destinationDirectory = ConfigurationManager.AppSettings["DestinationDirectory"];

                // Ensure the destination directory exists
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                // Get all XML files from the source directory
                string[] xmlFiles = Directory.GetFiles(sourceDirectory, "*.xml");

                if (xmlFiles.Length == 0)
                {
                    Console.WriteLine("No XML files found in the source directory.");
                    return;
                }

                foreach (string xmlFilePath in xmlFiles)
                {
                    try
                    {
                        // Get the file name without extension
                        string fileName = Path.GetFileNameWithoutExtension(xmlFilePath);

                        // Set the output Excel file path
                        string outputFilePath = Path.Combine(destinationDirectory, fileName + ".xlsx");

                        // Convert XML to Excel
                        var xmlDoc = XDocument.Load(xmlFilePath);
                        ConvertXmlToExcel(xmlDoc, outputFilePath);

                        Console.WriteLine($"Successfully converted '{xmlFilePath}' to Excel.");

                        // Optionally, delete the source XML file after conversion
                        File.Delete(xmlFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting file {xmlFilePath}: {ex.Message}");
                    }
                }

                Console.WriteLine("Conversion process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Method to convert XML to Excel
        public static void ConvertXmlToExcel(XDocument xmlDoc, string outputFilePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Get the root element and iterate through the child elements
                var root = xmlDoc.Root;
                if (root != null)
                {
                    var headers = root.Elements().First().Elements().Select(e => e.Name.LocalName).ToList();

                    // Write headers to the first row
                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                    }

                    // Write data to the subsequent rows
                    int row = 2;
                    foreach (var element in root.Elements())
                    {
                        int col = 1;
                        foreach (var value in element.Elements())
                        {
                            worksheet.Cell(row, col).Value = value.Value;
                            col++;
                        }
                        row++;
                    }
                }

                // Save the Excel file
                workbook.SaveAs(outputFilePath);
            }
        }
    }
}