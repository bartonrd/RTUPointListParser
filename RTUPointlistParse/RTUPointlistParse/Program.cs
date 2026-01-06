using UglyToad.PdfPig;
using ClosedXML.Excel;
using System.Text;
using System.Diagnostics;

namespace RTUPointlistParse
{
    public class Program
    {
        private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
        private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";

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
            // Note: In a real implementation with text-based PDFs, you would need to:
            // 1. Distinguish between Status and Analog data based on content or filename
            // 2. Parse different table structures for each type
            // For image-based PDFs (like the current example), this will result in empty collections
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
                        
                        // TODO: Implement logic to distinguish between Status and Analog data
                        // This could be based on:
                        // - Filename patterns (e.g., "sh1" for Status, "sh2" for Analog)
                        // - PDF content analysis (detecting specific header text)
                        // - Table structure differences
                        // For now, attempt to categorize based on filename
                        string fileName = Path.GetFileNameWithoutExtension(pdfFile).ToLower();
                        if (fileName.Contains("sh1") || fileName.Contains("status"))
                        {
                            allStatusRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} Status rows");
                        }
                        else if (fileName.Contains("sh2") || fileName.Contains("analog"))
                        {
                            allAnalogRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} Analog rows");
                        }
                        else
                        {
                            // If type is unknown, add to both (suboptimal but safe)
                            allStatusRows.AddRange(tableRows);
                            allAnalogRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} rows (type unknown)");
                        }
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
        /// Extract text content from a PDF file, using OCR if necessary
        /// </summary>
        public static string ExtractTextFromPdf(string filePath)
        {
            var sb = new StringBuilder();

            try
            {
                // First, try direct text extraction using PdfPig
                using (var document = PdfDocument.Open(filePath))
                {
                    foreach (var page in document.GetPages())
                    {
                        sb.AppendLine(page.Text);
                    }
                }

                // If no text was extracted, try OCR
                if (string.IsNullOrWhiteSpace(sb.ToString()))
                {
                    Console.WriteLine($"  No text found, attempting OCR...");
                    var ocrText = ExtractTextFromPdfWithOcr(filePath);
                    sb.Clear();
                    sb.Append(ocrText);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error extracting text: {ex.Message}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Check if a command-line tool is available in the system PATH
        /// </summary>
        private static bool IsToolAvailable(string toolName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = toolName,
                        Arguments = "--version",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                process.Start();
                process.WaitForExit();
                // Both exit code 0 (success) and 1 (some tools like pdftoppm) are acceptable
                // as they indicate the tool exists and responded
                return process.ExitCode == 0 || process.ExitCode == 1;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // Tool not found in PATH (common on Windows)
                return false;
            }
            catch (FileNotFoundException)
            {
                // Tool executable not found
                return false;
            }
            catch (Exception)
            {
                // Any other error means tool is not accessible
                return false;
            }
        }

        /// <summary>
        /// Display detailed diagnostic information for missing tesseract tool
        /// </summary>
        private static void DisplayTesseractDiagnostics()
        {
            Console.WriteLine("  Diagnostic information:");
            Console.WriteLine("    - Current PATH variable:");
            
            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (pathVar != null)
            {
                var paths = pathVar.Split(Path.PathSeparator);
                bool foundTesseract = false;
                
                foreach (var path in paths)
                {
                    if (path.Contains("Tesseract", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"      Found Tesseract in PATH: {path}");
                        foundTesseract = true;
                        
                        // Check if tesseract.exe exists in this path
                        var tesseractExe = Path.Combine(path, "tesseract.exe");
                        if (File.Exists(tesseractExe))
                        {
                            Console.WriteLine($"      ✓ tesseract.exe found at: {tesseractExe}");
                        }
                        else
                        {
                            Console.WriteLine($"      ✗ tesseract.exe NOT found at: {tesseractExe}");
                        }
                    }
                }
                
                if (!foundTesseract)
                {
                    Console.WriteLine("      No Tesseract directory found in PATH");
                }
            }
            
            // Check common installation locations on Windows
            if (OperatingSystem.IsWindows())
            {
                Console.WriteLine("    - Checking common installation locations:");
                var commonPaths = new[]
                {
                    @"C:\Program Files\Tesseract-OCR\tesseract.exe",
                    @"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                    @"C:\Tesseract-OCR\tesseract.exe"
                };
                
                foreach (var path in commonPaths)
                {
                    if (File.Exists(path))
                    {
                        Console.WriteLine($"      ✓ Found tesseract.exe at: {path}");
                        Console.WriteLine("      → You may need to restart your terminal/IDE for PATH changes to take effect");
                        Console.WriteLine("      → Or ensure the directory is in PATH, not just the parent directory");
                    }
                }
            }
        }

        /// <summary>
        /// Extract text from PDF using OCR (for image-based PDFs)
        /// </summary>
        private static string ExtractTextFromPdfWithOcr(string pdfPath)
        {
            try
            {
                Console.WriteLine($"  Performing OCR on PDF...");
                
                // Check if required tools are available
                if (!IsToolAvailable("pdftoppm"))
                {
                    Console.WriteLine("  ERROR: 'pdftoppm' not found. OCR requires poppler-utils to be installed.");
                    Console.WriteLine("  Installation instructions:");
                    Console.WriteLine("    Windows: Download from https://blog.alivate.com.au/poppler-windows/");
                    Console.WriteLine("             Extract and add the 'bin' folder to your system PATH");
                    Console.WriteLine("    Linux:   sudo apt-get install poppler-utils");
                    Console.WriteLine("    macOS:   brew install poppler");
                    Console.WriteLine("  ");
                    Console.WriteLine("  TROUBLESHOOTING:");
                    Console.WriteLine("    - After installing, restart your terminal/IDE/Command Prompt");
                    Console.WriteLine("    - Verify installation by running: pdftoppm -v");
                    return string.Empty;
                }

                if (!IsToolAvailable("tesseract"))
                {
                    Console.WriteLine("  ERROR: 'tesseract' not found. OCR requires Tesseract OCR to be installed.");
                    Console.WriteLine("  Installation instructions:");
                    Console.WriteLine("    Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki");
                    Console.WriteLine("             Install and ensure it's added to your system PATH");
                    Console.WriteLine("    Linux:   sudo apt-get install tesseract-ocr");
                    Console.WriteLine("    macOS:   brew install tesseract");
                    Console.WriteLine("  ");
                    Console.WriteLine("  TROUBLESHOOTING:");
                    Console.WriteLine("    - After installing, restart your terminal/IDE/Command Prompt");
                    Console.WriteLine("    - Verify installation by running: tesseract --version");
                    DisplayTesseractDiagnostics();
                    return string.Empty;
                }
                
                // Create a temporary directory for image files
                string tempDir = Path.Combine(Path.GetTempPath(), $"pdf_ocr_{Guid.NewGuid()}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Convert PDF pages to images using pdftoppm
                    var ppmProcess = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = "pdftoppm",
                            Arguments = $"-png \"{pdfPath}\" \"{Path.Combine(tempDir, "page")}\"",
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };

                    ppmProcess.Start();
                    ppmProcess.WaitForExit();

                    if (ppmProcess.ExitCode != 0)
                    {
                        var error = ppmProcess.StandardError.ReadToEnd();
                        Console.WriteLine($"  Error converting PDF to images: {error}");
                        return string.Empty;
                    }

                    // Get all generated image files
                    var imageFiles = Directory.GetFiles(tempDir, "*.png").OrderBy(f => f).ToArray();
                    
                    if (imageFiles.Length == 0)
                    {
                        Console.WriteLine($"  No images generated from PDF");
                        return string.Empty;
                    }

                    var sb = new StringBuilder();

                    // Perform OCR on each page image
                    foreach (var imageFile in imageFiles)
                    {
                        using (var tessProcess = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = "tesseract",
                                Arguments = $"\"{imageFile}\" stdout",
                                RedirectStandardOutput = true,
                                RedirectStandardError = true,
                                UseShellExecute = false,
                                CreateNoWindow = true
                            }
                        })
                        {
                            tessProcess.Start();
                            
                            // Read output before waiting to avoid deadlock
                            var text = tessProcess.StandardOutput.ReadToEnd();
                            var error = tessProcess.StandardError.ReadToEnd();
                            
                            tessProcess.WaitForExit();

                            // Only append text if tesseract succeeded
                            if (tessProcess.ExitCode == 0)
                            {
                                sb.AppendLine(text);
                            }
                            else
                            {
                                Console.WriteLine($"  Tesseract OCR error on {Path.GetFileName(imageFile)}: {error}");
                            }
                        }
                    }

                    Console.WriteLine($"  OCR completed on {imageFiles.Length} page(s)");
                    return sb.ToString();
                }
                finally
                {
                    // Clean up temporary files
                    try
                    {
                        Directory.Delete(tempDir, true);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  OCR extraction error: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Parse table data from extracted PDF text into structured rows
        /// </summary>
        public static List<TableRow> ParseTable(string pdfText)
        {
            var rows = new List<TableRow>();

            // Split text into lines
            var lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Detect if this is Status or Analog data based on content
            bool isStatusData = pdfText.Contains("STATUS STATE PAIR") || 
                                (pdfText.Contains("NORMAL") && pdfText.Contains("ALARM"));
            bool isAnalogData = pdfText.Contains("COEFFICIENT") || pdfText.Contains("OFFSET") || 
                                pdfText.Contains("FULL SCALE") || pdfText.Contains("SCALING");

            // Parse each line looking for table data rows
            // The table has a specific format with TAB DEC number at start of each row
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Skip header lines
                if (line.Contains("POINT NAME") && line.Contains("DNP"))
                    continue;
                if (line.Contains("DEC STATE"))
                    continue;
                if (line.Contains("PLOT BY:") || line.Contains("ISSUED FOR"))
                    continue;

                // Look for table data rows: start with number, followed by pipe or spaces, then data
                // Pattern: "NUMBER | POINT_NAME | ..." or "NUMBER spaces DATA..."
                var match = System.Text.RegularExpressions.Regex.Match(line, 
                    @"^(\d+)\s*[\|_\s]+(.+)");
                
                if (match.Success)
                {
                    var parsedRows = ParseTableLine(match, isStatusData, isAnalogData);
                    rows.AddRange(parsedRows);
                }
            }

            // Remove duplicates and sort by DNP INDEX
            // Also filter out invalid entries
            var uniqueRows = new Dictionary<int, TableRow>();
            foreach (var row in rows)
            {
                if (row.Columns.Count > 0)
                {
                    var dnpIndexStr = System.Text.RegularExpressions.Regex.Replace(
                        row.Columns[0].Trim(), @"[_\s]+$", "");
                    
                    if (int.TryParse(dnpIndexStr, out int dnpIndex))
                    {
                        // Filter out invalid DNP indices
                        // Valid range is typically 0-300 for point lists
                        if (dnpIndex < 0 || dnpIndex >= 500)
                            continue;
                        
                        // Filter out rows with no meaningful point name
                        if (row.Columns.Count < 3 || string.IsNullOrWhiteSpace(row.Columns[2]))
                            continue;
                        
                        // Filter out rows where point name is just special characters or very short
                        var pointName = row.Columns[2].Trim();
                        if (pointName.Length < 2 || pointName.All(c => !char.IsLetterOrDigit(c)))
                            continue;
                        
                        // Keep the row with more columns (more complete data)
                        if (!uniqueRows.ContainsKey(dnpIndex) || 
                            uniqueRows[dnpIndex].Columns.Count < row.Columns.Count)
                        {
                            uniqueRows[dnpIndex] = row;
                        }
                    }
                }
            }

            // Return sorted list
            return uniqueRows.OrderBy(kvp => kvp.Key).Select(kvp => kvp.Value).ToList();
        }

        /// <summary>
        /// Parse a table line that may contain one or two data entries (left and right columns)
        /// </summary>
        private static List<TableRow> ParseTableLine(System.Text.RegularExpressions.Match match, 
                                                      bool isStatusData, bool isAnalogData)
        {
            var result = new List<TableRow>();
            var fullLine = match.Value;
            
            // The line contains two entries side by side (left and right table columns)
            // Pattern: "NUMBER | DATA ... NUMBER[_|] DATA"
            // Split by finding positions where we have a number followed by | or __|
            // that appears after some content (not at the start)
            
            var regex = new System.Text.RegularExpressions.Regex(@"(\d+)\s*[_\s]*\|");
            var matches = regex.Matches(fullLine);
            
            if (matches.Count > 1)
            {
                // Multiple entries on this line
                for (int i = 0; i < matches.Count; i++)
                {
                    var startPos = matches[i].Index;
                    var endPos = (i < matches.Count - 1) ? matches[i + 1].Index : fullLine.Length;
                    var entry = fullLine.Substring(startPos, endPos - startPos).Trim();
                    
                    var columns = ParseDataEntry(entry, isStatusData, isAnalogData);
                    if (columns.Count > 0)
                    {
                        result.Add(new TableRow { Columns = columns });
                    }
                }
            }
            else if (matches.Count == 1)
            {
                // Single entry
                var columns = ParseDataEntry(fullLine, isStatusData, isAnalogData);
                if (columns.Count > 0)
                {
                    result.Add(new TableRow { Columns = columns });
                }
            }
            
            return result;
        }

        /// <summary>
        /// Parse a single data entry from the table
        /// </summary>
        private static List<string> ParseDataEntry(string entry, bool isStatusData, bool isAnalogData)
        {
            var columns = new List<string>();
            
            // Extract DNP INDEX (first number)
            var dnpMatch = System.Text.RegularExpressions.Regex.Match(entry, @"^(\d+)");
            if (!dnpMatch.Success)
                return columns;
            
            var dnpIndex = dnpMatch.Groups[1].Value;
            columns.Add(dnpIndex);
            
            // Remove DNP index and leading separator from entry
            var remainingData = entry.Substring(dnpMatch.Length).TrimStart('|', '_', ' ', '\t');
            
            if (isStatusData)
            {
                // Status format: [CONTROL_ADDR] POINT_NAME [NORMAL_STATE] [STATE1] [STATE0] ...
                // Split by pipe to separate major sections
                var parts = remainingData.Split('|').Select(p => p.Trim()).ToArray();
                
                if (parts.Length > 0)
                {
                    // First part: may contain point name and optional control address
                    var firstPart = parts[0];
                    var words = firstPart.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    string controlAddr = "";
                    string pointName = "";
                    
                    // Look for control address (a number after point name or at specific position)
                    // Common patterns: "POINT_NAME CTRL_NUM" or "CTRL_NUM POINT_NAME"
                    if (words.Length > 0)
                    {
                        // Try to find control address - it's usually a standalone number
                        int ctrlAddrIndex = -1;
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (int.TryParse(words[i], out int val) && val < 100 && i > 0)
                            {
                                ctrlAddrIndex = i;
                                break;
                            }
                        }
                        
                        if (ctrlAddrIndex > 0)
                        {
                            controlAddr = words[ctrlAddrIndex];
                            // Point name is words before control address
                            pointName = string.Join(" ", words.Take(ctrlAddrIndex));
                        }
                        else
                        {
                            // No control address found, entire first part is point name
                            pointName = firstPart;
                        }
                    }
                    
                    columns.Add(controlAddr);  // CONTROL ADDRESS (col 2)
                    columns.Add(pointName);    // POINT NAME (col 3)
                }
                
                // Extract state information from second part if present
                if (parts.Length > 1)
                {
                    var statePart = parts[1];
                    string normalState = "";
                    string state1 = "";
                    string state0 = "";
                    
                    // Look for state patterns: "CLOSE | OPEN" or "NORMAL | ALARM"
                    if (statePart.Contains("CLOSE") || statePart.Contains("OPEN"))
                    {
                        normalState = "1";
                        state1 = "CLOSE";
                        state0 = "OPEN";
                    }
                    else if (statePart.Contains("NORMAL") && statePart.Contains("ALARM"))
                    {
                        // Determine normal state based on order
                        if (statePart.IndexOf("NORMAL") < statePart.IndexOf("ALARM"))
                        {
                            normalState = "1";
                            state1 = "NORMAL";
                            state0 = "ALARM";
                        }
                        else
                        {
                            normalState = "0";
                            state1 = "ALARM";
                            state0 = "NORMAL";
                        }
                    }
                    else if (statePart.Contains("ALARM"))
                    {
                        normalState = "0";
                        state1 = "ALARM";
                        state0 = "NORMAL";
                    }
                    
                    columns.Add(normalState); // NORMAL STATE (col 4)
                    columns.Add(state1);      // 1_STATE (col 5)
                    columns.Add(state0);      // 0_STATE (col 6)
                }
            }
            else if (isAnalogData)
            {
                // Analog format: POINT_NAME [COEFFICIENT] [OFFSET] [VALUE] [UNIT] ...
                // For now, just extract point name from first part
                var parts = remainingData.Split('|').Select(p => p.Trim()).ToArray();
                
                if (parts.Length > 0)
                {
                    var pointName = parts[0].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    columns.Add(string.Join(" ", pointName)); // POINT NAME (col 2)
                }
            }

            return columns;
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
