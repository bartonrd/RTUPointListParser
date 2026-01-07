using UglyToad.PdfPig;
using ClosedXML.Excel;
using System.Text;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace RTUPointlistParse
{
    public class Program
    {
        private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
        private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";

        // Regex pattern to match data rows starting with a number
        private static readonly Regex DataRowPattern = new Regex(@"^\s*(\d+)\s*[|\[](.+)", RegexOptions.Compiled);

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
            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf").OrderBy(f => f).ToArray();
            Console.WriteLine($"Found {pdfFiles.Length} PDF file(s) to process");
            Console.WriteLine();

            if (pdfFiles.Length == 0)
            {
                Console.WriteLine("No PDF files found in the input folder.");
                return;
            }

            // Dictionary to collect all points: key = point number, value = point name
            var allPoints = new SortedDictionary<int, string>();

            // Process each PDF file
            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(pdfFile)}");
                    
                    // Extract text from all pages in the PDF
                    string pdfText = ExtractTextFromPdf(pdfFile);
                    
                    if (string.IsNullOrWhiteSpace(pdfText))
                    {
                        Console.WriteLine($"  Warning: No text extracted from PDF.");
                        continue;
                    }

                    // Extract point number and point name from the text
                    var points = ExtractPoints(pdfText);
                    
                    // Add to the global collection (avoiding duplicates)
                    foreach (var point in points)
                    {
                        if (!allPoints.ContainsKey(point.Key))
                        {
                            allPoints[point.Key] = point.Value;
                        }
                    }
                    
                    Console.WriteLine($"  Extracted {points.Count} points");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                }
            }

            // Generate Excel output
            string outputFileName = "Control_rtu837_DNP_pointlist_rev00.xlsx";
            string outputPath = Path.Combine(outputFolder, outputFileName);
            
            Console.WriteLine();
            Console.WriteLine($"Total points extracted: {allPoints.Count}");
            Console.WriteLine($"Generating Excel file: {outputFileName}");
            GenerateExcel(allPoints, outputPath);
            Console.WriteLine($"  Generated: {outputFileName}");

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
                return process.ExitCode == 0 || process.ExitCode == 1;
            }
            catch
            {
                return false;
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
                    return string.Empty;
                }

                if (!IsToolAvailable("tesseract"))
                {
                    Console.WriteLine("  ERROR: 'tesseract' not found. OCR requires Tesseract OCR to be installed.");
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
                        Console.WriteLine($"  Error converting PDF to images");
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
                            tessProcess.WaitForExit();

                            if (tessProcess.ExitCode == 0)
                            {
                                sb.AppendLine(text);
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
        /// Extract point numbers and point names from PDF text
        /// </summary>
        private static Dictionary<int, string> ExtractPoints(string pdfText)
        {
            var points = new Dictionary<int, string>();
            var lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var trimmedLine = line.Trim();

                // Skip header lines and metadata
                if (IsHeaderOrMetadata(trimmedLine))
                    continue;

                // Try to match a data row: starts with a number followed by | or [
                var match = DataRowPattern.Match(trimmedLine);
                if (match.Success)
                {
                    // Extract point number
                    if (!int.TryParse(match.Groups[1].Value, out int pointNumber))
                        continue;

                    // Filter out large numbers (likely reference/document numbers)
                    // Point numbers in the tables are typically small (1-300)
                    if (pointNumber > 500)
                        continue;

                    // Extract point name from the remainder
                    string remainder = match.Groups[2].Value.Trim();
                    string pointName = ExtractPointName(remainder);

                    // Only add valid point names
                    if (!string.IsNullOrWhiteSpace(pointName) && !IsSpareOrInvalid(pointName))
                    {
                        points[pointNumber] = pointName;
                    }
                }
            }

            return points;
        }

        /// <summary>
        /// Check if a line is a header or metadata (should be skipped)
        /// </summary>
        private static bool IsHeaderOrMetadata(string line)
        {
            // Skip common header patterns
            if (line.Contains("PLOT BY") || line.Contains("_PROJECTS") ||
                line.Contains(".dwg") || line.Contains("DIAG") ||
                line.Contains("POINT NAME") || line.Contains("DEC") ||
                line.Contains("STATE") || line.Contains("DSCRPT") ||
                line.Contains("RELAY NO") || line.Contains("INTERPOSING") ||
                line.Contains("COEFFICIENT") || line.Contains("OFFSET") ||
                line.StartsWith("i ") || line.StartsWith("a ") ||
                line.All(c => char.IsWhiteSpace(c) || c == '—' || c == '=' || c == '|'))
            {
                return true;
            }

            // Skip very short lines
            if (line.Length < 10)
                return true;

            return false;
        }

        /// <summary>
        /// Extract point name from the remainder of a data row
        /// </summary>
        private static string ExtractPointName(string text)
        {
            // Split by | to separate columns
            var parts = text.Split('|');
            if (parts.Length == 0)
                return string.Empty;

            // The first part typically contains the point name
            string firstPart = parts[0].Trim();

            // Remove any leading OCR artifacts
            firstPart = CleanOCRArtifacts(firstPart);

            // Extract the meaningful tokens (before control/state information)
            var tokens = firstPart.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            var nameTokens = new List<string>();
            int consecutiveSmallNumbers = 0;

            foreach (var token in tokens)
            {
                // Stop at state keywords
                if (token == "CLOSE" || token == "OPEN" || token == "NORMAL" || token == "ALARM" ||
                    token == "AUTO" || token == "SOLID" || token == "MANUAL" || token == "auto" ||
                    token.Contains("95-") || token.Contains("/AT") ||
                    token == "[" || token == "]" || token == "=" || token == "—")
                {
                    break;
                }

                // Clean the token
                string cleaned = CleanToken(token);
                if (string.IsNullOrWhiteSpace(cleaned) || cleaned.Length <= 1)
                {
                    continue;
                }

                // Check if this is a small standalone number (likely part of table structure, not name)
                if (int.TryParse(cleaned, out int num) && num < 100 && nameTokens.Count > 2)
                {
                    consecutiveSmallNumbers++;
                    // If we see multiple small numbers in a row, stop (likely entering data columns)
                    if (consecutiveSmallNumbers >= 2)
                        break;
                    continue;
                }
                else
                {
                    consecutiveSmallNumbers = 0;
                }

                nameTokens.Add(cleaned);

                // Stop after reasonable number of tokens
                if (nameTokens.Count >= 15)
                    break;
            }

            // Join and clean up the result
            string result = string.Join(" ", nameTokens).Trim();
            
            // Remove common OCR patterns at the end
            result = Regex.Replace(result, @"\s+[FI]\s*$", "");  // Trailing F or I
            result = Regex.Replace(result, @"\s+\d{1,2}\s*$", "");  // Trailing 1-2 digit numbers
            result = Regex.Replace(result, @"[F—\-~""]+\s*$", "");  // Trailing dashes, quotes, or F characters
            result = Regex.Replace(result, @"\s+[|\[\]]+\s*$", "");  // Trailing brackets
            result = Regex.Replace(result, @"\s+(lg|Tn|E\|)\s*$", "", RegexOptions.IgnoreCase);  // Common OCR trailing artifacts
            
            return result.Trim();
        }

        /// <summary>
        /// Clean OCR artifacts from a token
        /// </summary>
        private static string CleanToken(string token)
        {
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
                .Trim();

            // Remove leading OCR artifacts
            while (cleaned.Length > 1 && (cleaned[0] == 'l' || cleaned[0] == 'f' || cleaned[0] == 'I') &&
                   char.IsUpper(cleaned[1]))
            {
                cleaned = cleaned.Substring(1);
            }

            // Common OCR substitutions
            cleaned = cleaned
                .Replace("ftnyo", "INYO")
                .Replace("tnyo", "INYO")
                .Replace("NYO", "INYO")  // Common OCR error
                .Replace("GAS7AIR", "GAS/AIR")
                .Replace("T15KV", "115KV")
                .Replace("fi1S5KV", "115KV")
                .Replace("i1S5KV", "115KV")
                .Replace("lOXBOW", "OXBOW")
                .Replace("WWIXIE", "DIXIE")
                .Replace("fNO.", "NO.")
                .Replace("FNO.", "NO.")
                .Replace("INO.", "NO.")
                .Replace("FTRANS", "TRANS")
                .Replace("NHIBIT", "INHIBIT")  // Common OCR error
                .Replace("CBF", "CB")  // Trailing F artifact
                .Replace("coso", "COSO");  // Fix lowercase

            // Fix common patterns at start
            if (cleaned.StartsWith("N") && cleaned.Length > 2 && char.IsUpper(cleaned[1]))
            {
                // Check if it should be "I" (like NYO -> INYO)
                if (cleaned.StartsWith("NYO"))
                    cleaned = "I" + cleaned;
            }

            return cleaned;
        }

        /// <summary>
        /// Clean OCR artifacts from text
        /// </summary>
        private static string CleanOCRArtifacts(string text)
        {
            // Remove common OCR prefixes
            text = text.TrimStart('l', 'f', 'I', '/', '|', '[', ']');
            
            // Remove leading/trailing whitespace
            text = text.Trim();
            
            return text;
        }

        /// <summary>
        /// Check if a point name should be excluded (spare or invalid)
        /// </summary>
        private static bool IsSpareOrInvalid(string pointName)
        {
            if (string.IsNullOrWhiteSpace(pointName))
                return true;

            // Filter out SPARE points
            if (pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase) ||
                pointName.Contains("RESERVED", StringComparison.OrdinalIgnoreCase))
                return true;

            // Filter out very short names (likely OCR artifacts)
            if (pointName.Length <= 2)
                return true;

            // Filter out common OCR noise
            if (pointName == "I" || pointName == "F" || pointName == "J" || pointName == "L" ||
                pointName == "—" || pointName == "=" || pointName == "DI" || pointName == "or")
                return true;

            // Filter out metadata/reference lines
            if (pointName.Contains("LISTING") || pointName.Contains("CONSTRUCTION") ||
                pointName.Contains("ADDED POINT") || pointName.Contains("REPL'D RLY") ||
                pointName.Contains("REFERENCE DRAWINGS") || pointName.Contains("SAP") ||
                pointName.Contains("PLOT BY") || pointName.Contains("ISSUED FOR") ||
                pointName.Contains("REVISIONS") || pointName.Contains("DIG PT LISTING") ||
                pointName.Contains("ANALOG PT LISTING") || pointName.Contains("ED RTU/EPAC"))
                return true;

            return false;
        }

        /// <summary>
        /// Generate an Excel file with two columns: Point Number and Point Name
        /// </summary>
        private static void GenerateExcel(SortedDictionary<int, string> points, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                // Create a single sheet
                var worksheet = workbook.Worksheets.Add("Points");

                // Add header row
                worksheet.Cell(1, 1).Value = "Point Number";
                worksheet.Cell(1, 1).Style.Font.Bold = true;
                worksheet.Cell(1, 2).Value = "Point Name";
                worksheet.Cell(1, 2).Style.Font.Bold = true;

                // Add data rows
                int currentRow = 2;
                foreach (var point in points)
                {
                    worksheet.Cell(currentRow, 1).Value = point.Key;
                    worksheet.Cell(currentRow, 2).Value = point.Value;
                    currentRow++;
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(outputPath);
            }
        }
    }
}
