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

        // Constants for data parsing
        private const string DEFAULT_AOR_VALUE = "43";  // Default Area of Responsibility
        private const int MAX_POINT_NAME_TOKENS = 10;   // Maximum tokens to collect for point names
        private const int MAX_CONTROL_ADDRESS = 100;    // Maximum valid control address value
        private const string DEFAULT_NORMAL_STATE = "1";  // Default normal state value
        private const int POINT_NAME_COLUMN_INDEX = 2;  // Column index for point name (Status)
        private const int POINT_NAME_COLUMN_INDEX_ANALOG = 1;  // Column index for point name (Analog)
        
        // Constants for two-column layout detection
        private const int MIN_SECOND_COLUMN_INDEX = 80;  // Minimum row index indicating second column
        private const int MIN_CHAR_POSITION = 50;        // Minimum character position for second column split

        // Cached Regex patterns for better performance
        private static readonly System.Text.RegularExpressions.Regex DataRowPattern = 
            new System.Text.RegularExpressions.Regex(@"^\d+\s*[|\[]", 
                System.Text.RegularExpressions.RegexOptions.Compiled);
        private static readonly System.Text.RegularExpressions.Regex IndexExtractionPattern =
            new System.Text.RegularExpressions.Regex(@"^(\d+)\s*[|\[](.+)", 
                System.Text.RegularExpressions.RegexOptions.Compiled);
        private static readonly System.Text.RegularExpressions.Regex AlarmClassPattern =
            new System.Text.RegularExpressions.Regex(@"Class\s+(\d+)", 
                System.Text.RegularExpressions.RegexOptions.Compiled);
        private static readonly System.Text.RegularExpressions.Regex WhitespaceNormalizePattern =
            new System.Text.RegularExpressions.Regex(@"\s+", 
                System.Text.RegularExpressions.RegexOptions.Compiled);
        private static readonly System.Text.RegularExpressions.Regex TwoColumnSplitPattern =
            new System.Text.RegularExpressions.Regex(@"\s+(\d{2,3})\s*[_|]",
                System.Text.RegularExpressions.RegexOptions.Compiled);
        private static readonly System.Text.RegularExpressions.Regex TrailingSingleLetterPattern =
            new System.Text.RegularExpressions.Regex(@"\s+[A-Z]$",
                System.Text.RegularExpressions.RegexOptions.Compiled);
        
        // OCR artifact patterns
        private static readonly HashSet<string> OcrNoisePatterns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "I", "F", "J", "L", "—", "=", "DI", "or"
        };
        private static readonly char[] LeadingOcrArtifacts = new[] { 'l', 'f', 'I' };

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
                        // Determine type based on filename
                        string fileName = Path.GetFileNameWithoutExtension(pdfFile).ToLower();
                        
                        if (fileName.Contains("sh1") || fileName.Contains("status"))
                        {
                            // Parse as Status data
                            var tableRows = ParseStatusTable(pdfText);
                            allStatusRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} Status rows");
                        }
                        else if (fileName.Contains("sh2") || fileName.Contains("analog"))
                        {
                            // Parse as Analog data
                            var tableRows = ParseAnalogTable(pdfText);
                            allAnalogRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} Analog rows");
                        }
                        else
                        {
                            // Unknown type - try to parse as status
                            var tableRows = ParseStatusTable(pdfText);
                            allStatusRows.AddRange(tableRows);
                            Console.WriteLine($"  Extracted {tableRows.Count} rows (assumed Status)");
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
        /// Parse table data from extracted PDF text into structured rows for Status sheet
        /// Extracts only Point Number and Point Name columns
        /// </summary>
        public static List<TableRow> ParseStatusTable(string pdfText)
        {
            var rows = new List<TableRow>();
            var lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int pointNumber = 0;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var trimmedLine = line.Trim();

                // Skip metadata and header lines
                if (IsMetadataOrHeaderLine(trimmedLine))
                    continue;

                // Split two-column layout if present
                // PDFs may have two columns side-by-side that OCR reads as one line
                // Look for pattern like: "... 80 | POINT_NAME ..." which indicates second column start
                var columnLines = SplitTwoColumnLayout(trimmedLine);

                foreach (var columnLine in columnLines)
                {
                    if (string.IsNullOrWhiteSpace(columnLine))
                        continue;

                    // Check if this looks like a data row (starts with number followed by | or [)
                    if (DataRowPattern.IsMatch(columnLine))
                    {
                        // Parse this as a status data row - extract only Point Number and Point Name
                        var parsedRow = ParseSimpleDataRow(columnLine, pointNumber);
                        if (parsedRow != null)
                        {
                            string pointName = parsedRow.Columns.Count > 1 ? parsedRow.Columns[1] : "";
                            
                            // Filter out empty rows and rows where Point Name contains "Spare"
                            if (IsValidPointName(pointName))
                            {
                                rows.Add(parsedRow);
                                pointNumber++;
                            }
                        }
                    }
                }
            }

            return rows;
        }

        /// <summary>
        /// Parse table data from extracted PDF text into structured rows for Analog sheet
        /// Extracts only Point Number and Point Name columns
        /// </summary>
        public static List<TableRow> ParseAnalogTable(string pdfText)
        {
            var rows = new List<TableRow>();
            var lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int pointNumber = 0;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var trimmedLine = line.Trim();

                // Skip metadata and header lines
                if (IsMetadataOrHeaderLine(trimmedLine))
                    continue;

                // Split two-column layout if present
                var columnLines = SplitTwoColumnLayout(trimmedLine);

                foreach (var columnLine in columnLines)
                {
                    if (string.IsNullOrWhiteSpace(columnLine))
                        continue;

                    // Check if this looks like a data row
                    if (DataRowPattern.IsMatch(columnLine))
                    {
                        // Parse this as an analog data row - extract only Point Number and Point Name
                        var parsedRow = ParseSimpleDataRow(columnLine, pointNumber);
                        if (parsedRow != null)
                        {
                            string pointName = parsedRow.Columns.Count > 1 ? parsedRow.Columns[1] : "";
                            
                            // Filter out empty rows and rows where Point Name contains "Spare"
                            if (IsValidPointName(pointName))
                            {
                                rows.Add(parsedRow);
                                pointNumber++;
                            }
                        }
                    }
                }
            }

            return rows;
        }

        /// <summary>
        /// Backwards compatibility - parse as Status table
        /// </summary>
        public static List<TableRow> ParseTable(string pdfText)
        {
            return ParseStatusTable(pdfText);
        }

        /// <summary>
        /// Split a line that contains two-column layout data into separate column lines
        /// OCR may read two table columns side-by-side as a single line
        /// </summary>
        private static List<string> SplitTwoColumnLayout(string line)
        {
            var result = new List<string>();
            
            // Look for pattern indicating second column: a number (typically 80-300) followed by [_|] that appears mid-line
            var matches = TwoColumnSplitPattern.Matches(line);
            
            if (matches.Count == 0)
            {
                // No second column detected, return original line
                result.Add(line);
                return result;
            }
            
            // Find the best split point - look for indices >= MIN_SECOND_COLUMN_INDEX which typically indicate second column
            int bestSplitIndex = -1;
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (int.TryParse(match.Groups[1].Value, out int rowIndex))
                {
                    // Second column typically starts at index >= MIN_SECOND_COLUMN_INDEX
                    if (rowIndex >= MIN_SECOND_COLUMN_INDEX && match.Index > MIN_CHAR_POSITION)
                    {
                        bestSplitIndex = match.Index;
                        break;
                    }
                }
            }
            
            if (bestSplitIndex > 0)
            {
                // Split at this point
                string leftColumn = line.Substring(0, bestSplitIndex).Trim();
                string rightColumn = line.Substring(bestSplitIndex).Trim();
                
                if (!string.IsNullOrWhiteSpace(leftColumn))
                    result.Add(leftColumn);
                if (!string.IsNullOrWhiteSpace(rightColumn))
                    result.Add(rightColumn);
            }
            else
            {
                // No valid split point found, return original
                result.Add(line);
            }
            
            return result;
        }

        /// <summary>
        /// Check if a point name is valid (not empty and does not contain "Spare")
        /// Uses case-insensitive comparison to match variations like "SPARE", "Spare", "spare"
        /// </summary>
        /// <param name="pointName">The point name to validate</param>
        /// <returns>True if the point name is valid, false otherwise</returns>
        private static bool IsValidPointName(string pointName)
        {
            if (string.IsNullOrWhiteSpace(pointName))
                return false;
            
            // Filter out SPARE
            if (pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase))
                return false;
            
            // Filter out single characters or very short OCR artifacts
            if (pointName.Length <= 2)
                return false;
            
            // Filter out common OCR noise patterns
            if (pointName == "I" || pointName == "F" || pointName == "J" || pointName == "L" ||
                pointName == "—" || pointName == "=" || pointName == "DI" || pointName == "or")
                return false;
            
            // Filter out reference/metadata lines
            if (pointName.Contains("LISTING") || pointName.Contains("CONSTRUCTION") || 
                pointName.Contains("ADDED POINT") || pointName.Contains("SYSTEM") ||
                pointName.Contains("REFERENCE") || pointName.Contains("SAP") ||
                pointName.Contains("PLOT BY") || pointName.StartsWith("RESERVED FOR"))
                return false;
            
            return true;
        }

        /// <summary>
        /// Check if a line is metadata or header (should be skipped)
        /// </summary>
        private static bool IsMetadataOrHeaderLine(string line)
        {
            // Skip lines that are clearly metadata
            if (line.Contains("PLOT BY:") || line.Contains("_PROJECTS\\") ||
                line.Contains(".dwg") || line.Contains("DIAG") ||
                line.StartsWith("i ") || line.StartsWith("a ") ||
                (line.Contains("—") && line.Length < 20) ||
                (line.Contains("NOTE") && line.Contains("ADDED POINT")))
            {
                return true;
            }

            // Skip header rows (contain mostly column titles without data)
            if ((line.Contains("POINT NAME") && line.Contains("STATE")) ||
                (line.Contains("DEC") && line.Contains("DSCRPT")) ||
                (line.Contains("COEFFICIENT") && line.Contains("OFFSET")) ||
                line.Contains("INTERPOSNG") || line.Contains("RELAY NO."))
            {
                return true;
            }

            // Skip lines that are just separators or very short
            if (line.Length < 10 || line.All(c => char.IsWhiteSpace(c) || c == '—' || c == '=' || c == '|'))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Parse a data row from OCR text extracting only Point Number and Point Name
        /// </summary>
        private static TableRow? ParseSimpleDataRow(string line, int pointNumber)
        {
            try
            {
                // Pattern: NUMBER | POINT_NAME ... 
                var match = IndexExtractionPattern.Match(line);
                if (!match.Success)
                    return null;

                string remainder = match.Groups[2].Value;

                // Split by | to separate sections
                var sections = remainder.Split('|');
                if (sections.Length < 1)
                    return null;

                // First section contains the point name
                string firstSection = sections[0].Trim();

                // Extract point name (everything before certain keywords)
                string pointName = ExtractPointName(firstSection);
                if (string.IsNullOrWhiteSpace(pointName))
                    return null;

                // Build the row with only Point Number and Point Name
                var columns = new List<string>
                {
                    pointNumber.ToString(),  // Point Number
                    pointName                // Point Name
                };

                return new TableRow { Columns = columns };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Warning: Failed to parse row: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Parse a Status data row from OCR text
        /// Expected columns: TAB, CONTROL_ADDR, POINT_NAME, NORMAL_STATE, 1_STATE, 0_STATE, AOR, DOG_1, DOG_2, EMS_TP, VOLTAGE_BASE, ...
        /// </summary>
        private static TableRow? ParseStatusDataRow(string line, int tabIndex)
        {
            try
            {
                // Pattern: NUMBER | POINT_NAME ... CONTROL_INFO ... NORMAL_STATE | STATE ... 
                var match = IndexExtractionPattern.Match(line);
                if (!match.Success)
                    return null;

                string remainder = match.Groups[2].Value;

                // Split by | to separate sections
                var sections = remainder.Split('|');
                if (sections.Length < 2)
                    return null;

                // First section contains: POINT_NAME and possibly control address and state info
                string firstSection = sections[0].Trim();
                string secondSection = sections.Length > 1 ? sections[1].Trim() : "";

                // Extract point name (everything before certain keywords or numbers pattern)
                string pointName = ExtractPointName(firstSection);
                if (string.IsNullOrWhiteSpace(pointName))
                    return null;

                // Try to extract control address (small number after point name)
                string controlAddr = ExtractControlAddress(firstSection, pointName);

                // Extract state information (NORMAL_STATE, 1_STATE, 0_STATE)
                var (normalState, state1, state0) = ExtractStateInfo(firstSection, secondSection);

                // Build the row with available data
                var columns = new List<string>
                {
                    tabIndex.ToString(),           // TAB DEC DNP INDEX
                    controlAddr,                    // CONTROL ADDRESS
                    pointName,                      // POINT NAME
                    normalState,                    // NORMAL STATE
                    state1,                         // 1_STATE
                    state0,                         // 0_STATE
                    DEFAULT_AOR_VALUE,              // AOR (default)
                    ExtractAlarmClass(line, 1),    // DOG_1
                    ExtractAlarmClass(line, 2),    // DOG_2
                    "",                            // EMS TP NUMBER (not readily available)
                    ExtractVoltage(pointName)       // VOLTAGE BASE
                };

                return new TableRow { Columns = columns };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Warning: Failed to parse status row: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Parse an Analog data row from OCR text
        /// Expected columns: TAB, POINT_NAME, COEFFICIENT, OFFSET, VALUE, UNIT, LOW_LIMIT, HIGH_LIMIT, AOR, DOG_1, DOG_2, ...
        /// </summary>
        private static TableRow? ParseAnalogDataRow(string line, int tabIndex)
        {
            try
            {
                var match = IndexExtractionPattern.Match(line);
                if (!match.Success)
                    return null;

                string remainder = match.Groups[2].Value;
                var sections = remainder.Split('|');
                if (sections.Length < 1)
                    return null;

                string firstSection = sections[0].Trim();

                // Extract point name
                string pointName = ExtractPointName(firstSection);
                if (string.IsNullOrWhiteSpace(pointName))
                    return null;

                // Build the row with available data (many fields will be empty for OCR data)
                var columns = new List<string>
                {
                    tabIndex.ToString(),    // TAB DEC DNP INDEX
                    pointName,              // POINT NAME
                    "",                     // COEFFICIENT
                    "",                     // OFFSET
                    "",                     // VALUE
                    "",                     // UNIT
                    "",                     // LOW LIMIT
                    "",                     // HIGH LIMIT
                    DEFAULT_AOR_VALUE,      // AOR (default)
                    ExtractAlarmClass(line, 1),  // DOG_1
                    ExtractAlarmClass(line, 2),  // DOG_2
                };

                return new TableRow { Columns = columns };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Warning: Failed to parse analog row: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extract point name from the first part of the line
        /// </summary>
        private static string ExtractPointName(string text)
        {
            // Point names typically come first, before control info or state info
            // Common patterns: "NAME 115KV CB", "NAME SWITCH", etc.
            // Stop at: numbers followed by state keywords, certain characters like [, (

            var tokens = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            var nameTokens = new List<string>();
            bool hasSeenMainContent = false;
            bool justSawNo = false;
            bool justSawBank = false;

            foreach (var token in tokens)
            {
                // Stop collecting if we hit state keywords or control markers
                if (token == "CLOSE" || token == "OPEN" || token == "NORMAL" || token == "ALARM" ||
                    token == "AUTO" || token == "SOLID" || token == "MANUAL" || token == "auto" ||
                    token.Contains("95-") || token.Contains("/AT") || token == "[or" || token == "[ot" ||
                    token == "[pI" || token == "[oI" || token == "[dI" || token == "DI" || token.Contains("RK") ||
                    token == "[" || token == "]" || token == "=" || token == "—")
                {
                    break;
                }

                // Clean up OCR artifacts
                string cleaned = CleanOCRArtifacts(token);
                
                // Skip if cleaned token is empty or just punctuation
                if (string.IsNullOrWhiteSpace(cleaned) || cleaned == "—" || cleaned == "=" || cleaned.Length < 1)
                {
                    continue;
                }

                // Special case: allow numbers after "NO." or "BANK" (e.g., "NO. 1 BANK", "NO. 3 BANK")
                if (justSawNo || justSawBank)
                {
                    if (int.TryParse(cleaned, out int num) && num >= 0 && num <= 10)
                    {
                        nameTokens.Add(cleaned);
                        justSawNo = false;
                        justSawBank = cleaned == "BANK";
                        hasSeenMainContent = true;
                        continue;
                    }
                }

                // Skip standalone single-digit or two-digit numbers that appear after we've collected some content
                if (hasSeenMainContent && int.TryParse(cleaned, out int numVal))
                {
                    // If this is a small standalone number and we already have content, stop
                    if (cleaned.Length <= 2 && numVal < 200)
                    {
                        break;
                    }
                }

                // Skip tokens that are just numbers (unless part of a name like "NO. 1")
                if (int.TryParse(cleaned, out _) && !justSawNo && !justSawBank && 
                    !nameTokens.Contains("NO.") && !nameTokens.Contains("PLANT"))
                {
                    // Allow small numbers that might be part of names only if we don't have content yet
                    if (cleaned.Length <= 2 && !hasSeenMainContent)
                    {
                        nameTokens.Add(cleaned);
                        hasSeenMainContent = true;
                    }
                    else
                    {
                        break;  // Stop at standalone numbers
                    }
                }
                else
                {
                    nameTokens.Add(cleaned);
                    hasSeenMainContent = true;
                    justSawNo = (cleaned == "NO." || cleaned.EndsWith("NO."));
                    justSawBank = (cleaned == "BANK");
                }

                // Stop after collecting enough tokens (point names are usually 2-8 words)
                if (nameTokens.Count >= MAX_POINT_NAME_TOKENS)
                    break;
            }

            string result = string.Join(" ", nameTokens).Trim();
            
            // Final cleanup
            result = result.Replace("  ", " ");  // Remove double spaces
            result = WhitespaceNormalizePattern.Replace(result, " ");  // Normalize whitespace
            
            // Remove trailing single letters that are OCR artifacts (like "I", "F", "J")
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+[A-Z]$", "").Trim();
            
            return result;
        }

        /// <summary>
        /// Clean OCR artifacts from a token
        /// </summary>
        private static string CleanOCRArtifacts(string token)
        {
            // Common OCR artifacts
            string cleaned = token
                .Replace("[f", "")
                .Replace("||", "")
                .Replace("|", "")
                .Replace("[", "")
                .Replace("]", "")
                .Replace("(", "")
                .Replace(")", "")
                .Replace("{", "")
                .Replace("}", "")
                .Replace("_", " ")
                .Replace("  ", " ")
                .Trim();

            // Fix common OCR character confusions at the beginning
            // Remove leading lowercase 'l' or 'f' or 'I' that are likely artifacts
            while (cleaned.Length > 1 && (cleaned[0] == 'l' || cleaned[0] == 'f' || cleaned[0] == 'I') && 
                   char.IsUpper(cleaned[1]))
            {
                cleaned = cleaned.Substring(1);
            }

            if (cleaned.StartsWith("/") && cleaned.Length > 1)
            {
                // Leading slash is likely an OCR artifact
                cleaned = cleaned.Substring(1);
            }

            // Replace common character confusions
            cleaned = cleaned.Replace("I'", "1 ");
            cleaned = cleaned.Replace("Il", "11");
            cleaned = cleaned.Replace("Tn", "11");  // Common OCR confusion
            cleaned = cleaned.Replace("T15KV", "115KV");  // Specific OCR fix
            cleaned = cleaned.Replace("FTRANS", "TRANS");  // Remove leading F
            cleaned = cleaned.Replace("TS5KV", "115KV");  // Another OCR confusion
            cleaned = cleaned.Replace("IN15KV", "115KV");  // Another variation
            cleaned = cleaned.Replace("N15KV", "115KV");  // Another variation
            cleaned = cleaned.Replace("ftnyo", "INYO");  // Specific OCR fix
            cleaned = cleaned.Replace("WWIXIE", "DIXIE");  // Specific OCR fix
            cleaned = cleaned.Replace("FINO", "NO");  // Remove leading F
            cleaned = cleaned.Replace("FNO", "NO");  // Remove leading F
            cleaned = cleaned.Replace("INO", "NO");  // Remove leading I
            cleaned = cleaned.Replace("FCASA", "CASA");  // Remove leading F
            cleaned = cleaned.Replace("ICASA", "CASA");  // Remove leading I
            cleaned = cleaned.Replace("fi1S", "115");  // OCR confusion
            cleaned = cleaned.Replace("fF", "");  // Remove OCR artifact
            cleaned = cleaned.Replace("cD", "CD");  // Fix lowercase
            cleaned = cleaned.Replace("1155KV", "115KV");  // Fix OCR error
            cleaned = cleaned.Replace("NYO", "INYO");  // Fix missing I
            cleaned = cleaned.Replace("GAS7AIR", "GAS/AIR");  // Fix OCR
            cleaned = cleaned.Replace("CSF", "CS");  // Remove trailing F artifact
            cleaned = cleaned.Replace("CBF", "CB");  // Remove trailing F artifact
            
            return cleaned.Trim();
        }

        /// <summary>
        /// Extract control address (small number after point name)
        /// </summary>
        private static string ExtractControlAddress(string text, string pointName)
        {
            // Control address is typically a small number (0-99) that appears after the point name
            // and before state keywords
            try
            {
                int startPos = text.IndexOf(pointName);
                if (startPos < 0)
                    return "";

                string afterName = text.Substring(startPos + pointName.Length).Trim();
                var tokens = afterName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var token in tokens.Take(3)) // Check first few tokens after name
                {
                    // Look for a small number
                    if (int.TryParse(token, out int num) && num >= 0 && num < MAX_CONTROL_ADDRESS)
                    {
                        return token;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Warning: Failed to extract control address: {ex.Message}");
            }

            return "";
        }

        /// <summary>
        /// Extract state information (NORMAL/ALARM, CLOSE/OPEN, etc.)
        /// </summary>
        private static (string normal, string state1, string state0) ExtractStateInfo(string firstSection, string secondSection)
        {
            string normalState = DEFAULT_NORMAL_STATE;  // default
            string state1 = "";
            string state0 = "";

            // Look for state keywords
            if (firstSection.Contains("CLOSE") || secondSection.Contains("CLOSE"))
            {
                normalState = "1";
                state1 = "CLOSE";
                state0 = "OPEN";
            }
            else if (firstSection.Contains("OPEN") || secondSection.Contains("OPEN"))
            {
                normalState = "0";
                state1 = "CLOSE";
                state0 = "OPEN";
            }
            else if (firstSection.Contains("NORMAL"))
            {
                normalState = "1";
                state1 = "NORMAL";
                state0 = "ALARM";
            }
            else if (firstSection.Contains("ALARM"))
            {
                normalState = "0";
                state1 = "NORMAL";
                state0 = "ALARM";
            }
            else if (firstSection.Contains("AUTO") || firstSection.Contains("auto"))
            {
                normalState = "1";
                state1 = "AUTO";
                state0 = "SOLID";
            }

            return (normalState, state1, state0);
        }

        /// <summary>
        /// Extract alarm class (Class 1, Class 2, etc.)
        /// </summary>
        private static string ExtractAlarmClass(string line, int classNumber)
        {
            // Look for patterns like "Class 1", "Class 2", etc.
            var match = AlarmClassPattern.Match(line);
            if (match.Success && classNumber == 1)
            {
                return $"Class {match.Groups[1].Value}";
            }

            // Look for second class if exists
            var matches = AlarmClassPattern.Matches(line);
            if (matches.Count >= classNumber)
            {
                return $"Class {matches[classNumber - 1].Groups[1].Value}";
            }

            return "";
        }

        /// <summary>
        /// Extract voltage base from point name (115KV, 55KV, 0KV, etc.)
        /// </summary>
        private static string ExtractVoltage(string pointName)
        {
            if (pointName.Contains("115KV") || pointName.Contains("115 KV"))
                return "115KV";
            if (pointName.Contains("55KV") || pointName.Contains("55 KV"))
                return "55KV";
            if (pointName.Contains("12KV") || pointName.Contains("12 KV"))
                return "12KV";

            return "0KV";  // default
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
            // Add simple header
            worksheet.Cell(1, 1).Value = "Point Number";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "Point Name";
            worksheet.Cell(1, 2).Style.Font.Bold = true;

            // Add data rows
            int currentRow = 2;
            foreach (var row in rows)
            {
                for (int i = 0; i < Math.Min(row.Columns.Count, 2); i++)
                {
                    worksheet.Cell(currentRow, i + 1).Value = row.Columns[i];
                }
                currentRow++;
            }
        }

        private static void CreateAnalogSheet(IXLWorksheet worksheet, List<TableRow> rows)
        {
            // Add simple header
            worksheet.Cell(1, 1).Value = "Point Number";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "Point Name";
            worksheet.Cell(1, 2).Style.Font.Bold = true;

            // Add data rows
            int currentRow = 2;
            foreach (var row in rows)
            {
                for (int i = 0; i < Math.Min(row.Columns.Count, 2); i++)
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
