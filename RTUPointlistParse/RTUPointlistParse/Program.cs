using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace RTUPointlistParse
{
    /// <summary>
    /// Minimal OCR + Geometry Algorithm for parsing RTU Point Lists from PDFs
    /// </summary>
    public class Program
    {
        // Constants for detection
        private const int DPI = 300;
        private const double HEADER_Y_GAP_TOLERANCE = 20.0; // pixels for stacked headers
        private const double ROW_Y_TOLERANCE = 10.0;  // pixels for grouping words into rows
        private const int COL1_EXPAND = 15;  // expand column 1 band
        private const int COL2_EXPAND = 100;  // expand column 2 band (wider to capture full names)
        
        private static readonly Regex NumericPattern = new Regex(@"^\d+$", RegexOptions.Compiled);
        
        // Default paths
        private const string DefaultInputFolder = "ExamplePointlists/Example1/Input";
        private const string DefaultOutputFolder = "ExamplePointlists/Example1/TestOutput";
        
        public static void Main(string[] args)
        {
            string inputFolder = args.Length > 0 ? args[0] : DefaultInputFolder;
            string outputFolder = args.Length > 1 ? args[1] : DefaultOutputFolder;
            
            Console.WriteLine("RTU Point List Parser (OCR + Geometry)");
            Console.WriteLine("=======================================");
            Console.WriteLine($"Input folder: {inputFolder}");
            Console.WriteLine($"Output folder: {outputFolder}");
            Console.WriteLine();
            
            // Validate folders
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Error: Input folder does not exist: {inputFolder}");
                return;
            }
            
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            
            // Get all PDF files
            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
            Console.WriteLine($"Found {pdfFiles.Length} PDF file(s)\n");
            
            if (pdfFiles.Length == 0)
            {
                Console.WriteLine("No PDF files found.");
                return;
            }
            
            var logBuilder = new StringBuilder();
            logBuilder.AppendLine("RTU Point List Parser Log");
            logBuilder.AppendLine("=========================");
            logBuilder.AppendLine($"Processed at: {DateTime.Now}");
            logBuilder.AppendLine();
            
            // Collect all points from all PDFs
            var allPoints = new List<PointData>();
            
            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(pdfFile)}");
                    logBuilder.AppendLine($"File: {Path.GetFileName(pdfFile)}");
                    
                    var points = ProcessPdf(pdfFile, logBuilder);
                    allPoints.AddRange(points);
                    
                    Console.WriteLine($"  Extracted {points.Count} points");
                    logBuilder.AppendLine($"  Extracted: {points.Count} points");
                    logBuilder.AppendLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error: {ex.Message}");
                    logBuilder.AppendLine($"  ERROR: {ex.Message}");
                    logBuilder.AppendLine();
                }
            }
            
            // Sort by Point Number
            allPoints = allPoints.OrderBy(p => p.PointNumber).ToList();
            
            // Validate continuity
            logBuilder.AppendLine("Validation:");
            logBuilder.AppendLine($"  Total points: {allPoints.Count}");
            
            if (allPoints.Count > 0)
            {
                var gaps = new List<int>();
                var duplicates = new Dictionary<int, int>();
                
                for (int i = 0; i < allPoints.Count; i++)
                {
                    int expected = i + 1;
                    int actual = allPoints[i].PointNumber;
                    
                    if (actual != expected)
                    {
                        if (actual > expected)
                        {
                            for (int gap = expected; gap < actual; gap++)
                            {
                                gaps.Add(gap);
                            }
                        }
                    }
                    
                    // Check for duplicates
                    if (i > 0 && allPoints[i].PointNumber == allPoints[i-1].PointNumber)
                    {
                        if (duplicates.ContainsKey(actual))
                            duplicates[actual]++;
                        else
                            duplicates[actual] = 2;
                    }
                }
                
                if (gaps.Count > 0)
                {
                    logBuilder.AppendLine($"  GAPS: Missing point numbers: {string.Join(", ", gaps.Take(20))}");
                    if (gaps.Count > 20)
                        logBuilder.AppendLine($"    ... and {gaps.Count - 20} more");
                }
                else
                {
                    logBuilder.AppendLine("  No gaps detected");
                }
                
                if (duplicates.Count > 0)
                {
                    logBuilder.AppendLine($"  DUPLICATES: {string.Join(", ", duplicates.Select(d => $"{d.Key}({d.Value}x)"))}");
                }
                else
                {
                    logBuilder.AppendLine("  No duplicates detected");
                }
                
                logBuilder.AppendLine($"  Point range: {allPoints[0].PointNumber} to {allPoints[allPoints.Count-1].PointNumber}");
            }
            
            // Write Excel output
            string excelPath = Path.Combine(outputFolder, "PointList.xlsx");
            WriteExcel(allPoints, excelPath);
            Console.WriteLine($"\nGenerated: {Path.GetFileName(excelPath)}");
            logBuilder.AppendLine($"\nOutput: {Path.GetFileName(excelPath)}");
            
            // Write log file
            string logPath = Path.Combine(outputFolder, "PointList.log.txt");
            File.WriteAllText(logPath, logBuilder.ToString());
            Console.WriteLine($"Generated: {Path.GetFileName(logPath)}");
            
            Console.WriteLine("\nProcessing complete.");
        }
        
        private static List<PointData> ProcessPdf(string pdfPath, StringBuilder log)
        {
            var allPoints = new List<PointData>();
            
            // Get PDF page count
            using var pdfDoc = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
            int pageCount = pdfDoc.NumberOfPages;
            
            for (int pageNum = 1; pageNum <= pageCount; pageNum++)
            {
                Console.WriteLine($"  Page {pageNum}/{pageCount}");
                
                // Step 1: Render PDF page to image
                string imagePath = RenderPdfPageToImage(pdfPath, pageNum);
                if (imagePath == null)
                {
                    log.AppendLine($"    Page {pageNum}: Failed to render");
                    continue;
                }
                
                try
                {
                    // Step 2: OCR with bounding boxes
                    var words = OcrImageWithBoundingBoxes(imagePath);
                    log.AppendLine($"    Page {pageNum}: OCR extracted {words.Count} words");
                    
                    // Step 3: Detect table headers
                    var tables = DetectTableHeaders(words);
                    log.AppendLine($"    Page {pageNum}: Detected {tables.Count} table(s)");
                    
                    // Log table positions for debugging
                    foreach (var table in tables)
                    {
                        log.AppendLine($"      Table: PointNum@X={table.PointNumberHeaderX}, PointName@X={table.PointNameHeaderX}, HeaderY={table.HeaderY}");
                    }
                    
                    // Step 4-6: Extract points from each table
                    foreach (var table in tables)
                    {
                        var points = ExtractPointsFromTable(words, table);
                        allPoints.AddRange(points);
                    }
                }
                finally
                {
                    // Clean up temp image
                    try { File.Delete(imagePath); } catch { }
                }
            }
            
            return allPoints;
        }
        
        private static string? RenderPdfPageToImage(string pdfPath, int pageNum)
        {
            try
            {
                string tempDir = Path.GetTempPath();
                string outputPrefix = Path.Combine(tempDir, $"page_{Guid.NewGuid()}");
                
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "pdftoppm",
                        Arguments = $"-png -r {DPI} -f {pageNum} -l {pageNum} \"{pdfPath}\" \"{outputPrefix}\"",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                process.WaitForExit();
                
                if (process.ExitCode != 0)
                {
                    return null;
                }
                
                // pdftoppm creates files like: prefix-1.png
                string expectedFile = $"{outputPrefix}-{pageNum}.png";
                if (File.Exists(expectedFile))
                    return expectedFile;
                
                // Sometimes it uses page 1 for single page
                expectedFile = $"{outputPrefix}-1.png";
                if (File.Exists(expectedFile))
                    return expectedFile;
                
                return null;
            }
            catch
            {
                return null;
            }
        }
        
        private static List<OcrWord> OcrImageWithBoundingBoxes(string imagePath)
        {
            var words = new List<OcrWord>();
            
            try
            {
                // Use tesseract with hocr output to get bounding boxes
                string hocrPath = Path.Combine(Path.GetTempPath(), $"hocr_{Guid.NewGuid()}");
                
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "tesseract",
                        Arguments = $"\"{imagePath}\" \"{hocrPath}\" hocr",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                process.WaitForExit();
                
                string hocrFile = hocrPath + ".hocr";
                if (!File.Exists(hocrFile))
                {
                    return words;
                }
                
                try
                {
                    // Parse HOCR XML
                    var doc = XDocument.Load(hocrFile);
                    var wordElements = doc.Descendants()
                        .Where(e => e.Attribute("class")?.Value == "ocrx_word");
                    
                    foreach (var wordElem in wordElements)
                    {
                        string text = wordElem.Value.Trim();
                        if (string.IsNullOrWhiteSpace(text))
                            continue;
                        
                        string? title = wordElem.Attribute("title")?.Value;
                        if (title == null)
                            continue;
                        
                        // Parse bbox from title: "bbox 100 200 150 220"
                        var bboxMatch = Regex.Match(title, @"bbox (\d+) (\d+) (\d+) (\d+)");
                        if (bboxMatch.Success)
                        {
                            int x1 = int.Parse(bboxMatch.Groups[1].Value);
                            int y1 = int.Parse(bboxMatch.Groups[2].Value);
                            int x2 = int.Parse(bboxMatch.Groups[3].Value);
                            int y2 = int.Parse(bboxMatch.Groups[4].Value);
                            
                            words.Add(new OcrWord
                            {
                                Text = text,
                                X = x1,
                                Y = y1,
                                Width = x2 - x1,
                                Height = y2 - y1
                            });
                        }
                    }
                }
                finally
                {
                    try { File.Delete(hocrFile); } catch { }
                }
            }
            catch
            {
                // Fallback: plain OCR without bounding boxes - not ideal but better than nothing
            }
            
            return words;
        }
        
        private static List<TableInfo> DetectTableHeaders(List<OcrWord> words)
        {
            var tables = new List<TableInfo>();
            
            // Find all header-related words
            var pointWords = words.Where(w => 
                w.Text.Equals("POINT", StringComparison.OrdinalIgnoreCase)).ToList();
            
            var numberWords = words.Where(w =>
                w.Text.Equals("NUMBER", StringComparison.OrdinalIgnoreCase)).ToList();
                
            var nameWords = words.Where(w =>
                w.Text.Equals("NAME", StringComparison.OrdinalIgnoreCase)).ToList();
            
            // Strategy: Find "POINT" and "NUMBER" that are close together (same row or stacked)
            // forming a "POINT NUMBER" header column
            foreach (var numberWord in numberWords)
            {
                // Find a POINT word near this NUMBER
                var nearbyPoint = pointWords.FirstOrDefault(p =>
                    Math.Abs(p.Y - numberWord.Y) < 100 &&  // Same general row or stacked
                    Math.Abs(p.X - numberWord.X) < 800);   // Reasonably close horizontally
                
                if (nearbyPoint == null)
                    continue;
                
                // This forms the Point Number column header
                // Now find "POINT NAME" to the right
                int headerY = Math.Min(nearbyPoint.Y, numberWord.Y);
                int headerX = Math.Min(nearbyPoint.X, numberWord.X);
                int headerRightX = Math.Max(nearbyPoint.X + nearbyPoint.Width, numberWord.X + numberWord.Width);
                
                // Look for "POINT" to the right on a similar row
                var rightPointWords = pointWords.Where(p =>
                    p != nearbyPoint &&
                    p.X > headerRightX + 500 &&  // Well to the right
                    Math.Abs(p.Y - headerY) < 100).ToList();
                
                foreach (var rpw in rightPointWords)
                {
                    // Look for NAME near this POINT
                    var nearbyName = nameWords.FirstOrDefault(n =>
                        Math.Abs(n.Y - rpw.Y) < 100 &&
                        Math.Abs(n.X - rpw.X) < 800);
                    
                    if (nearbyName != null)
                    {
                        // Create table
                        tables.Add(new TableInfo
                        {
                            PointNumberHeaderX = headerX,
                            PointNumberHeaderWidth = headerRightX - headerX,
                            PointNameHeaderX = Math.Min(rpw.X, nearbyName.X),
                            PointNameHeaderWidth = Math.Max(rpw.X + rpw.Width, nearbyName.X + nearbyName.Width) - Math.Min(rpw.X, nearbyName.X),
                            HeaderY = Math.Max(headerY, Math.Max(numberWord.Y, nearbyPoint.Y))
                        });
                        break; // Only one table per POINT-NUMBER pair
                    }
                }
            }
            
            return tables;
        }
        
        private static List<PointData> ExtractPointsFromTable(List<OcrWord> words, TableInfo table)
        {
            var points = new List<PointData>();
            
            // Define column bands more precisely
            int col1MinX = table.PointNumberHeaderX - COL1_EXPAND;
            int col1MaxX = table.PointNumberHeaderX + table.PointNumberHeaderWidth + COL1_EXPAND;
            int col2MinX = table.PointNameHeaderX - COL2_EXPAND;
            int col2MaxX = table.PointNameHeaderX + table.PointNameHeaderWidth + (COL2_EXPAND * 3); // Wider on the right
            
            // Get words below the header
            var dataWords = words.Where(w => w.Y > table.HeaderY + 20).ToList();
            
            // Cluster by Y coordinate (group into rows)
            var rows = ClusterWordsByY(dataWords, ROW_Y_TOLERANCE);
            
            foreach (var row in rows)
            {
                // Extract Point Number: first numeric word in col1 band
                var numberWords = row.Where(w =>
                    w.X >= col1MinX && w.X <= col1MaxX &&
                    NumericPattern.IsMatch(w.Text)).ToList();
                
                if (numberWords.Count == 0)
                    continue;
                
                var numberWord = numberWords.OrderBy(w => w.X).First();
                if (!int.TryParse(numberWord.Text, out int pointNumber))
                    continue;
                
                // Extract Point Name: words within the Point Name column band
                var nameWords = row.Where(w =>
                    w.X >= col2MinX && w.X <= col2MaxX).OrderBy(w => w.X).ToList();
                
                if (nameWords.Count == 0)
                    continue;
                
                // Concatenate words intelligently
                var nameTokens = new List<string>();
                int lastX = col2MinX;
                bool foundMainName = false;
                
                foreach (var word in nameWords)
                {
                    // Skip if too far from previous word (indicates column break)
                    if (foundMainName && (word.X - lastX > 300))
                        break;
                    
                    // Skip pure numeric after we have name content (likely next column data)
                    if (foundMainName && NumericPattern.IsMatch(word.Text) && word.Text.Length <= 2)
                        break;
                    
                    // Add word
                    nameTokens.Add(word.Text);
                    lastX = word.X + word.Width;
                    foundMainName = true;
                    
                    // Limit total tokens
                    if (nameTokens.Count >= 15)
                        break;
                }
                
                string pointName = string.Join(" ", nameTokens).Trim();
                
                if (string.IsNullOrWhiteSpace(pointName))
                    continue;
                
                // Filter out noise
                if (pointName.Length <= 2 || 
                    pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase))
                    continue;
                
                // Clean up common OCR artifacts
                pointName = CleanPointName(pointName);
                
                if (string.IsNullOrWhiteSpace(pointName) || pointName.Length <= 2)
                    continue;
                
                points.Add(new PointData
                {
                    PointNumber = pointNumber,
                    PointName = pointName
                });
            }
            
            return points;
        }
        
        private static string CleanPointName(string name)
        {
            // Remove leading OCR artifacts
            name = name.TrimStart('|', 'f', 'I', 'l', '[', ']');
            
            // Remove trailing punctuation artifacts
            name = name.TrimEnd('|', '[', ']', '\\');
            
            // Fix common OCR errors
            name = name.Replace("||", "");
            name = name.Replace("|f", "");
            name = name.Replace("[GATE", "");
            name = name.Replace("fI", "");
            
            // Normalize whitespace
            name = Regex.Replace(name, @"\s+", " ");
            
            return name.Trim();
        }
        
        private static List<List<OcrWord>> ClusterWordsByY(List<OcrWord> words, double tolerance)
        {
            if (words.Count == 0)
                return new List<List<OcrWord>>();
            
            // Sort by Y
            var sorted = words.OrderBy(w => w.Y).ToList();
            var clusters = new List<List<OcrWord>>();
            var currentCluster = new List<OcrWord> { sorted[0] };
            
            for (int i = 1; i < sorted.Count; i++)
            {
                if (Math.Abs(sorted[i].Y - currentCluster[0].Y) <= tolerance)
                {
                    currentCluster.Add(sorted[i]);
                }
                else
                {
                    clusters.Add(currentCluster);
                    currentCluster = new List<OcrWord> { sorted[i] };
                }
            }
            clusters.Add(currentCluster);
            
            return clusters;
        }
        
        private static void WriteExcel(List<PointData> points, string outputPath)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Points");
            
            // Headers
            sheet.Cell(1, 1).Value = "Point Number";
            sheet.Cell(1, 1).Style.Font.Bold = true;
            sheet.Cell(1, 2).Value = "Point Name";
            sheet.Cell(1, 2).Style.Font.Bold = true;
            
            // Data
            for (int i = 0; i < points.Count; i++)
            {
                sheet.Cell(i + 2, 1).Value = points[i].PointNumber;
                sheet.Cell(i + 2, 2).Value = points[i].PointName;
            }
            
            workbook.SaveAs(outputPath);
        }
    }
    
    public class OcrWord
    {
        public string Text { get; set; } = "";
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
    
    public class TableInfo
    {
        public int PointNumberHeaderX { get; set; }
        public int PointNumberHeaderWidth { get; set; }
        public int PointNameHeaderX { get; set; }
        public int PointNameHeaderWidth { get; set; }
        public int HeaderY { get; set; }
    }
    
    public class PointData
    {
        public int PointNumber { get; set; }
        public string PointName { get; set; } = "";
    }
}
