using UglyToad.PdfPig;
using ClosedXML.Excel;
using System.Text;

namespace RTUPointlistParse
{
    public class Program
    {
        private const string DefaultInputFolder = "ExamplePointlists/Example1/Input";
        private const string DefaultOutputFolder = "ExamplePointlists/Example1/TestOutput";

        public static void Main(string[] args)
        {
            // Parse command-line arguments
            string inputFolder = args.Length > 0 ? args[0] : DefaultInputFolder;
            string outputFolder = args.Length > 1 ? args[1] : DefaultOutputFolder;

            Console.WriteLine("RTU Point List Parser");
            Console.WriteLine("=====================");
            Console.WriteLine($"Input folder: {inputFolder}");
            Console.WriteLine($"Output folder: {outputFolder}");
            Console.WriteLine();

            // Validate input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Error: Input folder does not exist: {inputFolder}");
                return;
            }

            // Create output folder if it doesn't exist
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
                Console.WriteLine($"Created output folder: {outputFolder}");
            }

            // Enumerate all PDF files in the input folder
            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
            Console.WriteLine($"Found {pdfFiles.Length} PDF file(s) to process");
            Console.WriteLine();

            if (pdfFiles.Length == 0)
            {
                Console.WriteLine("No PDF files found in the input folder.");
                return;
            }

            // Collect all table rows from all PDFs
            var allStatusRows = new List<TableRow>();
            var allAnalogRows = new List<TableRow>();

            // Process each PDF file
            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(pdfFile)}");
                    
                    // Extract text from PDF
                    string pdfText = ExtractTextFromPdf(pdfFile);
                    
                    if (string.IsNullOrWhiteSpace(pdfText))
                    {
                        Console.WriteLine($"  Warning: No text extracted from PDF. The PDF may be image-based.");
                    }
                    else
                    {
                        // Parse the table from the extracted text
                        var tableRows = ParseTable(pdfText);
                        
                        // Determine if this is Status or Analog data based on filename or content
                        // For now, add to both collections
                        allStatusRows.AddRange(tableRows);
                        allAnalogRows.AddRange(tableRows);
                        
                        Console.WriteLine($"  Extracted {tableRows.Count} rows");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                }
            }

            // Generate a single combined Excel file
            // Use a generic name or derive from folder/first file
            string outputFileName = "Control_rtu837_DNP_pointlist_rev00.xlsx";
            string outputPath = Path.Combine(outputFolder, outputFileName);
            
            Console.WriteLine();
            Console.WriteLine($"Generating combined Excel file: {outputFileName}");
            GenerateExcel(allStatusRows, allAnalogRows, outputPath);
            Console.WriteLine($"  Generated: {outputFileName}");

            // Compare generated files with expected output if available
            string expectedOutputFolder = Path.Combine(
                Path.GetDirectoryName(inputFolder) ?? "",
                "Expected Output"
            );

            if (Directory.Exists(expectedOutputFolder))
            {
                Console.WriteLine();
                Console.WriteLine("Comparing with expected output...");
                Console.WriteLine("=================================");
                CompareOutputFiles(outputFolder, expectedOutputFolder);
            }

            Console.WriteLine();
            Console.WriteLine("Processing complete.");
        }

        /// <summary>
        /// Extract text content from a PDF file
        /// </summary>
        public static string ExtractTextFromPdf(string filePath)
        {
            var sb = new StringBuilder();

            try
            {
                using (var document = PdfDocument.Open(filePath))
                {
                    foreach (var page in document.GetPages())
                    {
                        sb.AppendLine(page.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error extracting text: {ex.Message}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Parse table data from extracted PDF text into structured rows
        /// </summary>
        public static List<TableRow> ParseTable(string pdfText)
        {
            var rows = new List<TableRow>();

            // Split text into lines
            var lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Parse the line into columns (simple split by whitespace for now)
                var columns = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                if (columns.Count > 0)
                {
                    rows.Add(new TableRow { Columns = columns });
                }
            }

            return rows;
        }

        /// <summary>
        /// Generate an Excel file from parsed table rows
        /// </summary>
        public static void GenerateExcel(List<TableRow> statusRows, List<TableRow> analogRows, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                // Create Status sheet
                var statusSheet = workbook.Worksheets.Add("Status");
                CreateStatusSheet(statusSheet, statusRows);

                // Create Analog sheet
                var analogSheet = workbook.Worksheets.Add("Analog");
                CreateAnalogSheet(analogSheet, analogRows);

                workbook.SaveAs(outputPath);
            }
        }

        private static void CreateStatusSheet(IXLWorksheet worksheet, List<TableRow> rows)
        {
            // Add header
            worksheet.Cell(1, 1).Value = "CONTRL_D DNP Status Point List";
            worksheet.Cell(1, 1).Style.Font.Bold = true;

            // Add metadata rows (simplified version)
            int currentRow = 3;
            worksheet.Cell(currentRow, 1).Value = "LOCATION: ";
            worksheet.Cell(currentRow, 5).Value = "RTU/DEVICE TYPE: ";
            worksheet.Cell(currentRow, 10).Value = "STA DC VOLTAGE: ";
            worksheet.Cell(currentRow, 23).Value = "NOTE: ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "DATE:";
            worksheet.Cell(currentRow, 5).Value = "EMS DEVICE NUM: ";
            worksheet.Cell(currentRow, 10).Value = "POINTLIST REVISION: ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "DEVICE NAME: ";
            worksheet.Cell(currentRow, 5).Value = "RTU/SAS DNP ADDRESS: ";
            worksheet.Cell(currentRow, 10).Value = "A SYSTEM:  ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "TDBU SAP: ";
            worksheet.Cell(currentRow, 5).Value = "PSC ENGINEER:  ";
            worksheet.Cell(currentRow, 10).Value = "SWITCHING CENTER:  ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "PSC SAP: ";
            worksheet.Cell(currentRow, 5).Value = "CRQ NUMBER:  ";
            worksheet.Cell(currentRow, 10).Value = "BES ASSET:  ";
            worksheet.Cell(currentRow, 23).Value = "TESTING HISTORY";

            // Add column headers
            currentRow += 2;
            worksheet.Cell(currentRow, 2).Value = "CONTROL ADDRESS ";
            worksheet.Cell(currentRow, 4).Value = "STATUS STATE PAIR INFO ";
            worksheet.Cell(currentRow, 7).Value = "ALARMS  ";
            worksheet.Cell(currentRow, 12).Value = "CROSS REFERENCE EXISTING EMS DATA ";
            worksheet.Cell(currentRow, 16).Value = "TAB-1 BASED ";
            worksheet.Cell(currentRow, 17).Value = "IED INFORMATION ";

            currentRow += 2;
            var headers = new[] {
                "TAB DEC DNP INDEX", "0 BASED CONTROL ADDRESS", "POINT NAME                    ",
                "NORMAL STATE", "1_STATE", "0_STATE", "AOR", " DOG_1 /3  ", "  DOG_2 /4   ",
                "EMS TP NUMBER", "VOLTAGE BASE", "EXISTING DEVICE NAME", "EXISTING POINT NAME",
                "EXISTING TAB NUM", "ITEM  ", "CONTROL  ADDRESS", "LAN     (CARD_PORT)",
                "IED ADDRESS", "I/O_REGISTER       DNP_INDEX        ", "PLC_MAPPING   OBJECT_NAME    "
            };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(currentRow, i + 1).Value = headers[i];
                worksheet.Cell(currentRow, i + 1).Style.Font.Bold = true;
            }

            // Add data rows from parsed table
            currentRow++;
            foreach (var row in rows)
            {
                for (int i = 0; i < Math.Min(row.Columns.Count, headers.Length); i++)
                {
                    worksheet.Cell(currentRow, i + 1).Value = row.Columns[i];
                }
                currentRow++;
            }
        }

        private static void CreateAnalogSheet(IXLWorksheet worksheet, List<TableRow> rows)
        {
            // Add header
            worksheet.Cell(1, 1).Value = "CONTRL_D  DNP Analog Point List";
            worksheet.Cell(1, 1).Style.Font.Bold = true;

            // Add metadata rows (simplified version)
            int currentRow = 3;
            worksheet.Cell(currentRow, 1).Value = "LOCATION:  ";
            worksheet.Cell(currentRow, 5).Value = "RTU/DEVICE MODEL: ";
            worksheet.Cell(currentRow, 10).Value = "STA DC VOLTAGE: ";
            worksheet.Cell(currentRow, 18).Value = "NOTE: ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "DATE: ";
            worksheet.Cell(currentRow, 5).Value = "EMS DEVICE NUM: ";
            worksheet.Cell(currentRow, 10).Value = "POINTLIST REVISION: ";
            worksheet.Cell(currentRow, 18).Value = "All fullscale and limits are true values";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "DEVICE NAME: ";
            worksheet.Cell(currentRow, 5).Value = "RTU/SAS ADDRESS: ";
            worksheet.Cell(currentRow, 10).Value = "A SYSTEM: ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "TDBU SAP: ";
            worksheet.Cell(currentRow, 5).Value = "PSC ENGINEER:  ";
            worksheet.Cell(currentRow, 10).Value = "SWITCHING CENTER: ";

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "PSC SAP: ";
            worksheet.Cell(currentRow, 5).Value = "PSC TECHENICIAN:  ";
            worksheet.Cell(currentRow, 10).Value = "TESTMAN: ";

            // Add column group headers
            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "EMS DATABASE INFORMATION ";
            worksheet.Cell(currentRow, 14).Value = "CROSS REFERENCE INFORMATION ";
            worksheet.Cell(currentRow, 17).Value = "FIELD INFORMATION ";

            currentRow++;
            worksheet.Cell(currentRow, 3).Value = "SCALING ";
            worksheet.Cell(currentRow, 5).Value = "FULL SCALE ";
            worksheet.Cell(currentRow, 7).Value = "LIMITS ";
            worksheet.Cell(currentRow, 9).Value = "ALARMS ";
            worksheet.Cell(currentRow, 14).Value = "EXISTING EMS DATA ";
            worksheet.Cell(currentRow, 17).Value = "IED  INFORMATION ";

            currentRow++;
            var headers = new[] {
                "TAB DEC DNP INDEX", "POINT NAME", "COEFFICIENT", "OFFSET", "VALUE", "UNIT",
                "LOW LIMIT", "HIGH LIMIT", "       AOR        ", "       DOG_1/3        ",
                "    DOG_2/4     ", "EMS_TP NUMBER", "VOLTAGE BASE", "EXISTING DEVICE NAME",
                "EXISTING POINT NAME", "EXISTING TAB NUM", "ITEM", "LAN_CARD-PORT",
                "IED_ADDRESS", "I/O_REGISTER_or DNP_INDEX"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(currentRow, i + 1).Value = headers[i];
                worksheet.Cell(currentRow, i + 1).Style.Font.Bold = true;
            }

            // Add data rows from parsed table
            currentRow++;
            foreach (var row in rows)
            {
                for (int i = 0; i < Math.Min(row.Columns.Count, headers.Length); i++)
                {
                    worksheet.Cell(currentRow, i + 1).Value = row.Columns[i];
                }
                currentRow++;
            }
        }

        private static void CompareOutputFiles(string generatedFolder, string expectedFolder)
        {
            var generatedFiles = Directory.GetFiles(generatedFolder, "*.xlsx");

            foreach (var generatedFile in generatedFiles)
            {
                string fileName = Path.GetFileName(generatedFile);
                string expectedFile = Path.Combine(expectedFolder, fileName);

                if (!File.Exists(expectedFile))
                {
                    Console.WriteLine($"{fileName}: No expected output file found for comparison");
                    continue;
                }

                bool match = CompareExcelFiles(generatedFile, expectedFile);
                Console.WriteLine($"{fileName}: {(match ? "Match" : "Differences detected")}");
            }
        }

        /// <summary>
        /// Compare two Excel files and return true if they match
        /// </summary>
        public static bool CompareExcelFiles(string generatedFile, string expectedFile)
        {
            try
            {
                using (var generatedWorkbook = new XLWorkbook(generatedFile))
                using (var expectedWorkbook = new XLWorkbook(expectedFile))
                {
                    // Compare number of worksheets
                    if (generatedWorkbook.Worksheets.Count != expectedWorkbook.Worksheets.Count)
                    {
                        Console.WriteLine($"  - Different number of worksheets");
                        return false;
                    }

                    bool allMatch = true;

                    foreach (var expectedSheet in expectedWorkbook.Worksheets)
                    {
                        var generatedSheet = generatedWorkbook.Worksheets.FirstOrDefault(w => w.Name == expectedSheet.Name);

                        if (generatedSheet == null)
                        {
                            Console.WriteLine($"  - Missing worksheet: {expectedSheet.Name}");
                            allMatch = false;
                            continue;
                        }

                        // Compare used ranges
                        var expectedRange = expectedSheet.RangeUsed();
                        var generatedRange = generatedSheet.RangeUsed();

                        if (expectedRange == null && generatedRange == null)
                            continue;

                        if (expectedRange == null || generatedRange == null)
                        {
                            Console.WriteLine($"  - {expectedSheet.Name}: Different data presence");
                            allMatch = false;
                            continue;
                        }

                        // Compare dimensions
                        if (expectedRange.RowCount() != generatedRange.RowCount() ||
                            expectedRange.ColumnCount() != generatedRange.ColumnCount())
                        {
                            Console.WriteLine($"  - {expectedSheet.Name}: Different dimensions " +
                                $"(Expected: {expectedRange.RowCount()}x{expectedRange.ColumnCount()}, " +
                                $"Generated: {generatedRange.RowCount()}x{generatedRange.ColumnCount()})");
                            allMatch = false;
                            continue;
                        }

                        // Sample comparison of cell values (first 20 rows)
                        int rowsToCheck = Math.Min(20, expectedRange.RowCount());
                        int colsToCheck = expectedRange.ColumnCount();

                        for (int r = 1; r <= rowsToCheck; r++)
                        {
                            for (int c = 1; c <= colsToCheck; c++)
                            {
                                var expectedValue = expectedRange.Cell(r, c).GetValue<string>();
                                var generatedValue = generatedRange.Cell(r, c).GetValue<string>();

                                if (expectedValue != generatedValue)
                                {
                                    Console.WriteLine($"  - {expectedSheet.Name}: Cell mismatch at R{r}C{c} " +
                                        $"(Expected: '{expectedValue}', Generated: '{generatedValue}')");
                                    allMatch = false;
                                }
                            }
                        }
                    }

                    return allMatch;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error comparing files: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Represents a row of table data with multiple columns
    /// </summary>
    public class TableRow
    {
        public List<string> Columns { get; set; } = new List<string>();
    }
}
