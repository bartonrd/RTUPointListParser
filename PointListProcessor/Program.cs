using System;
using System.IO;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace PointListProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string inputFolder;
            string outputFolder;

            // Parse command-line arguments
            if (args.Length >= 2 && args.Contains("--input") && args.Contains("--output"))
            {
                int inputIndex = Array.IndexOf(args, "--input");
                int outputIndex = Array.IndexOf(args, "--output");
                
                // Validate that both flags have corresponding values
                if (inputIndex + 1 >= args.Length || outputIndex + 1 >= args.Length)
                {
                    Console.WriteLine("Error: --input and --output flags require values.");
                    Console.WriteLine("Usage: PointListProcessor --input <path> --output <path>");
                    return;
                }
                
                inputFolder = args[inputIndex + 1];
                outputFolder = args[outputIndex + 1];
            }
            else if (args.Length >= 2)
            {
                // Assume first arg is input, second is output (without flags)
                inputFolder = args[0];
                outputFolder = args[1];
            }
            else
            {
                // Prompt user for input and output folders
                Console.WriteLine("Enter input folder path (or press Enter for default 'Input'):");
                inputFolder = Console.ReadLine() ?? "";
                if (string.IsNullOrWhiteSpace(inputFolder))
                {
                    inputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Input");
                }

                Console.WriteLine("Enter output folder path (or press Enter for default 'TestOutput'):");
                outputFolder = Console.ReadLine() ?? "";
                if (string.IsNullOrWhiteSpace(outputFolder))
                {
                    outputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestOutput");
                }
            }

            // Validate and create directories
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Error: Input folder '{inputFolder}' does not exist.");
                return;
            }

            Directory.CreateDirectory(outputFolder);
            Console.WriteLine($"Input folder: {inputFolder}");
            Console.WriteLine($"Output folder: {outputFolder}");

            // Process PDF files
            var inputFiles = Directory.GetFiles(inputFolder, "*.pdf");
            
            if (inputFiles.Length == 0)
            {
                Console.WriteLine("No PDF files found in the input folder.");
                return;
            }

            Console.WriteLine($"Found {inputFiles.Length} PDF file(s) to process.");

            foreach (var file in inputFiles)
            {
                try
                {
                    Console.WriteLine($"\nProcessing: {Path.GetFileName(file)}");
                    string content = ExtractTextFromPdf(file);
                    var parsedData = ParsePointList(content);
                    string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(file) + "_output.xlsx");
                    GenerateOutput(parsedData, outputFile);
                    Console.WriteLine($"  Generated: {Path.GetFileName(outputFile)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error processing {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Console.WriteLine("\nProcessing complete!");
        }

        static string ExtractTextFromPdf(string filePath)
        {
            using (var pdfReader = new PdfReader(filePath))
            using (var pdfDocument = new PdfDocument(pdfReader))
            {
                var text = "";
                for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
                {
                    var page = pdfDocument.GetPage(i);
                    var strategy = new SimpleTextExtractionStrategy();
                    text += PdfTextExtractor.GetTextFromPage(page, strategy) + "\n";
                }
                return text;
            }
        }

        static List<PointData> ParsePointList(string content)
        {
            var points = new List<PointData>();
            
            // Split content into lines
            var lines = content.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            
            // Basic parsing logic - this is a simplified example that extracts text
            // TODO: Customize this regex pattern based on the actual point list format
            // Current pattern looks for: word word rest_of_line
            // Real-world formats may require more sophisticated parsing
            foreach (var line in lines)
            {
                // Skip header lines or empty lines
                if (string.IsNullOrWhiteSpace(line) || line.Contains("Point List") || line.Contains("Page"))
                    continue;

                // Example regex pattern to extract point information
                // This pattern is intentionally simple and should be customized for production use
                var match = Regex.Match(line, @"(\w+)\s+(\w+)\s+(.+)");
                if (match.Success)
                {
                    points.Add(new PointData
                    {
                        PointName = match.Groups[1].Value,
                        PointType = match.Groups[2].Value,
                        Description = match.Groups[3].Value
                    });
                }
            }

            return points;
        }

        static void GenerateOutput(List<PointData> points, string outputFile)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Point List");
                
                // Add headers
                worksheet.Cells[1, 1].Value = "Point Name";
                worksheet.Cells[1, 2].Value = "Point Type";
                worksheet.Cells[1, 3].Value = "Description";
                
                // Style headers
                using (var range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Add data
                for (int i = 0; i < points.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = points[i].PointName;
                    worksheet.Cells[i + 2, 2].Value = points[i].PointType;
                    worksheet.Cells[i + 2, 3].Value = points[i].Description;
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                // Save the file
                var fileInfo = new FileInfo(outputFile);
                package.SaveAs(fileInfo);
            }
        }
    }

    class PointData
    {
        public string PointName { get; set; } = "";
        public string PointType { get; set; } = "";
        public string Description { get; set; } = "";
    }
}
