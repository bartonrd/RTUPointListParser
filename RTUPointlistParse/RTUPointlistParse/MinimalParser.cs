using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace RTUPointlistParse
{
    /// <summary>
    /// Minimal algorithm: OCR + Geometry based table extraction
    /// </summary>
    public class MinimalParser
    {
        private const int DPI = 300; // Rendering DPI
        private const int HEADER_Y_GAP_TOLERANCE = 12; // Max Y gap between stacked headers
        private const int ROW_Y_TOLERANCE = 50; // Y tolerance for row clustering (increased for headers)
        private const int COL1_X_EXPAND = 15; // Expand X range for Point Number column
        private const int COL2_X_EXPAND = 200; // Expand X range for Point Name column (wide to capture full names)

        /// <summary>
        /// Process all PDF files and generate PointList.xlsx and log
        /// </summary>
        public static void ProcessPDFs(string inputFolder, string outputFolder)
        {
            var logBuilder = new StringBuilder();
            var allPoints = new List<PointData>();
            
            // Get all PDF files
            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
            logBuilder.AppendLine($"Processing {pdfFiles.Length} PDF file(s)");
            logBuilder.AppendLine();

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    logBuilder.AppendLine($"Processing: {Path.GetFileName(pdfFile)}");
                    var points = ProcessSinglePDF(pdfFile, logBuilder);
                    allPoints.AddRange(points);
                    logBuilder.AppendLine($"  Extracted {points.Count} point(s)");
                }
                catch (Exception ex)
                {
                    logBuilder.AppendLine($"  ERROR: {ex.Message}");
                }
                logBuilder.AppendLine();
            }

            // Sort by point number
            allPoints.Sort((a, b) => a.PointNumber.CompareTo(b.PointNumber));

            // Validate continuity
            logBuilder.AppendLine("Validation:");
            ValidateContinuity(allPoints, logBuilder);
            logBuilder.AppendLine();

            // Write Excel output
            string excelPath = Path.Combine(outputFolder, "PointList.xlsx");
            WriteExcelOutput(allPoints, excelPath);
            logBuilder.AppendLine($"Written: {excelPath}");

            // Write log file
            string logPath = Path.Combine(outputFolder, "PointList.log.txt");
            File.WriteAllText(logPath, logBuilder.ToString());
            logBuilder.AppendLine($"Written: {logPath}");

            Console.WriteLine(logBuilder.ToString());
        }

        /// <summary>
        /// Process a single PDF file
        /// </summary>
        private static List<PointData> ProcessSinglePDF(string pdfFile, StringBuilder logBuilder)
        {
            var allPoints = new List<PointData>();

            // Step 1: Render PDF pages to images
            var imageFiles = RenderPDFToImages(pdfFile);
            
            try
            {
                // Step 2: OCR each page with bounding boxes
                foreach (var imagePath in imageFiles)
                {
                    var words = ExtractWordsWithBoundingBoxes(imagePath);
                    
                    // Step 3: Detect table headers on this page
                    var tables = DetectTableHeaders(words);
                    
                    // Step 4: Extract rows from each table
                    foreach (var table in tables)
                    {
                        var points = ExtractTableRows(words, table);
                        allPoints.AddRange(points);
                    }
                }
            }
            finally
            {
                // Cleanup temp image files
                foreach (var img in imageFiles)
                {
                    try { File.Delete(img); } catch { }
                }
            }

            return allPoints;
        }

        /// <summary>
        /// Render PDF to images using pdftoppm
        /// </summary>
        private static List<string> RenderPDFToImages(string pdfFile)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), $"pdf_render_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "pdftoppm",
                    Arguments = $"-png -r {DPI} \"{pdfFile}\" \"{Path.Combine(tempDir, "page")}\"",
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new Exception($"pdftoppm failed: {process.StandardError.ReadToEnd()}");
            }

            return Directory.GetFiles(tempDir, "*.png").OrderBy(f => f).ToList();
        }

        /// <summary>
        /// Extract words with bounding boxes using Tesseract hOCR output
        /// </summary>
        private static List<Word> ExtractWordsWithBoundingBoxes(string imagePath)
        {
            var words = new List<Word>();
            string hocrFile = Path.Combine(Path.GetTempPath(), $"hocr_{Guid.NewGuid()}");

            try
            {
                // Run tesseract with hocr output
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "tesseract",
                        Arguments = $"\"{imagePath}\" \"{hocrFile}\" hocr",
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new Exception($"tesseract failed: {process.StandardError.ReadToEnd()}");
                }

                // Parse hOCR output
                string hocrPath = hocrFile + ".hocr";
                words = ParseHOCR(hocrPath);
                
                // Cleanup
                try { File.Delete(hocrPath); } catch { }
            }
            catch (Exception ex)
            {
                throw new Exception($"OCR extraction failed: {ex.Message}");
            }

            return words;
        }

        /// <summary>
        /// Parse hOCR HTML to extract words with bounding boxes
        /// </summary>
        private static List<Word> ParseHOCR(string hocrPath)
        {
            var words = new List<Word>();
            
            try
            {
                var doc = XDocument.Load(hocrPath);
                var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

                // Find all word elements (class="ocrx_word")
                var wordElements = doc.Descendants()
                    .Where(e => e.Attribute("class")?.Value.Contains("ocrx_word") == true);

                foreach (var elem in wordElements)
                {
                    var title = elem.Attribute("title")?.Value;
                    var text = elem.Value.Trim();

                    if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(title))
                        continue;

                    // Parse bbox from title: "bbox x0 y0 x1 y1"
                    var match = Regex.Match(title, @"bbox (\d+) (\d+) (\d+) (\d+)");
                    if (match.Success)
                    {
                        words.Add(new Word
                        {
                            Text = text,
                            X0 = int.Parse(match.Groups[1].Value),
                            Y0 = int.Parse(match.Groups[2].Value),
                            X1 = int.Parse(match.Groups[3].Value),
                            Y1 = int.Parse(match.Groups[4].Value)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to parse hOCR: {ex.Message}");
            }

            return words;
        }

        /// <summary>
        /// Detect table headers: "POINT" and "NUMBER" (horizontal or stacked) and "POINT NAME" or "NAME"
        /// </summary>
        private static List<TableHeader> DetectTableHeaders(List<Word> words)
        {
            var tables = new List<TableHeader>();

            // Strategy 1: Look for stacked "POINT" over "NUMBER"
            var pointWords = words.Where(w => 
                w.Text.Equals("POINT", StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var pointWord in pointWords)
            {
                // Look for "NUMBER" below POINT (overlapping X, small Y gap)
                var numberWord = words.FirstOrDefault(w =>
                    w.Text.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                    w.Y0 > pointWord.Y1 &&
                    w.Y0 - pointWord.Y1 <= HEADER_Y_GAP_TOLERANCE &&
                    Math.Max(w.X0, pointWord.X0) < Math.Min(w.X1, pointWord.X1)); // X overlap

                if (numberWord != null)
                {
                    // Found stacked POINT/NUMBER - now find POINT NAME to the right
                    var pointNameWord = words.FirstOrDefault(w =>
                        (w.Text.Equals("POINT", StringComparison.OrdinalIgnoreCase) ||
                         w.Text.StartsWith("POINT", StringComparison.OrdinalIgnoreCase)) &&
                        w.X0 > numberWord.X1 &&
                        Math.Abs(w.Y0 - pointWord.Y0) <= ROW_Y_TOLERANCE); // Same row band

                    // Also look for "NAME" alone
                    if (pointNameWord == null)
                    {
                        pointNameWord = words.FirstOrDefault(w =>
                            w.Text.Equals("NAME", StringComparison.OrdinalIgnoreCase) &&
                            w.X0 > numberWord.X1 &&
                            Math.Abs(w.Y0 - pointWord.Y0) <= ROW_Y_TOLERANCE);
                    }

                    if (pointNameWord != null)
                    {
                        tables.Add(new TableHeader
                        {
                            PointNumberX0 = Math.Min(pointWord.X0, numberWord.X0) - COL1_X_EXPAND,
                            PointNumberX1 = Math.Max(pointWord.X1, numberWord.X1) + COL1_X_EXPAND,
                            PointNameX0 = pointNameWord.X0 - COL2_X_EXPAND,
                            PointNameX1 = pointNameWord.X1 + COL2_X_EXPAND,
                            HeaderY1 = Math.Max(pointWord.Y1, numberWord.Y1),
                        });
                    }
                }
            }

            // Strategy 2: Look for horizontal arrangement "NUMBER" ... "POINT" ... "NAME"
            if (tables.Count == 0)
            {
                var numberWords = words.Where(w => 
                    w.Text.Equals("NUMBER", StringComparison.OrdinalIgnoreCase)).ToList();

                foreach (var numberWord in numberWords)
                {
                    // Look for POINT to the right in the same row
                    var pointWord = words.FirstOrDefault(w =>
                        w.Text.Equals("POINT", StringComparison.OrdinalIgnoreCase) &&
                        w.X0 > numberWord.X1 &&
                        Math.Abs(w.Y0 - numberWord.Y0) <= ROW_Y_TOLERANCE);

                    if (pointWord != null)
                    {
                        // Look for NAME to the right of POINT
                        var nameWord = words.FirstOrDefault(w =>
                            w.Text.Equals("NAME", StringComparison.OrdinalIgnoreCase) &&
                            w.X0 > pointWord.X1 &&
                            Math.Abs(w.Y0 - pointWord.Y0) <= ROW_Y_TOLERANCE);

                        if (nameWord != null)
                        {
                            tables.Add(new TableHeader
                            {
                                PointNumberX0 = numberWord.X0 - COL1_X_EXPAND,
                                PointNumberX1 = numberWord.X1 + COL1_X_EXPAND,
                                PointNameX0 = pointWord.X0 - COL2_X_EXPAND,  // Start from POINT, not NAME
                                PointNameX1 = nameWord.X1 + COL2_X_EXPAND * 2,  // Extend further right
                                HeaderY1 = Math.Max(numberWord.Y1, Math.Max(pointWord.Y1, nameWord.Y1)),
                            });
                        }
                    }
                }
            }

            return tables;
        }

        /// <summary>
        /// Extract rows from a table
        /// </summary>
        private static List<PointData> ExtractTableRows(List<Word> words, TableHeader table)
        {
            var points = new List<PointData>();

            // Filter words below header
            var dataWords = words.Where(w => w.Y0 > table.HeaderY1).ToList();

            // Cluster rows by Y coordinate
            var rows = ClusterRowsByY(dataWords);

            foreach (var row in rows)
            {
                // Extract Point Number (numeric token in Col1 band)
                var pointNumWord = row.FirstOrDefault(w =>
                    w.X0 >= table.PointNumberX0 && w.X0 <= table.PointNumberX1 &&
                    int.TryParse(w.Text, out _));

                if (pointNumWord == null)
                    continue; // Skip rows without numeric point number

                if (!int.TryParse(pointNumWord.Text, out int pointNumber))
                    continue;

                // Extract Point Name - only words strictly within the Point Name X band
                var nameWords = row.Where(w =>
                    w.X0 >= table.PointNameX0 && w.X0 <= table.PointNameX1)
                    .OrderBy(w => w.X0)
                    .ToList();

                string pointName = string.Join(" ", nameWords.Select(w => w.Text)).Trim();

                // Clean up the point name
                pointName = CleanPointName(pointName);

                if (string.IsNullOrWhiteSpace(pointName))
                    continue; // Skip rows with empty point name

                points.Add(new PointData
                {
                    PointNumber = pointNumber,
                    PointName = pointName
                });
            }

            return points;
        }

        /// <summary>
        /// Clean up point name by removing OCR artifacts and extra delimiters
        /// </summary>
        private static string CleanPointName(string pointName)
        {
            // Remove leading/trailing delimiters
            pointName = pointName.Trim('|', '=', '—', '_', ' ', '[', ']', '(', ')');
            
            // Remove sequences of vertical bars
            pointName = Regex.Replace(pointName, @"\|+", " ");
            
            // Remove sequences of equals or dashes
            pointName = Regex.Replace(pointName, @"[=—_]+", " ");
            
            // Replace multiple spaces with single space
            pointName = Regex.Replace(pointName, @"\s+", " ");

            return pointName.Trim();
        }

        /// <summary>
        /// Cluster words into rows by Y coordinate
        /// </summary>
        private static List<List<Word>> ClusterRowsByY(List<Word> words)
        {
            var rows = new List<List<Word>>();
            var sorted = words.OrderBy(w => w.Y0).ThenBy(w => w.X0).ToList();

            foreach (var word in sorted)
            {
                // Find existing row with similar Y
                var row = rows.FirstOrDefault(r =>
                    Math.Abs(r[0].Y0 - word.Y0) <= ROW_Y_TOLERANCE);

                if (row != null)
                {
                    row.Add(word);
                }
                else
                {
                    rows.Add(new List<Word> { word });
                }
            }

            // Sort each row left to right
            foreach (var row in rows)
            {
                row.Sort((a, b) => a.X0.CompareTo(b.X0));
            }

            // Sort rows top to bottom
            rows.Sort((a, b) => a[0].Y0.CompareTo(b[0].Y0));

            return rows;
        }

        /// <summary>
        /// Validate continuity and log gaps/duplicates
        /// </summary>
        private static void ValidateContinuity(List<PointData> points, StringBuilder logBuilder)
        {
            if (points.Count == 0)
            {
                logBuilder.AppendLine("  No points extracted");
                return;
            }

            logBuilder.AppendLine($"  Total points: {points.Count}");
            logBuilder.AppendLine($"  Range: {points[0].PointNumber} to {points[points.Count - 1].PointNumber}");

            // Check for duplicates
            var duplicates = points.GroupBy(p => p.PointNumber)
                .Where(g => g.Count() > 1)
                .ToList();

            if (duplicates.Any())
            {
                logBuilder.AppendLine("  DUPLICATES:");
                foreach (var dup in duplicates)
                {
                    logBuilder.AppendLine($"    Point {dup.Key}: {dup.Count()} occurrences");
                }
            }

            // Check for gaps
            var gaps = new List<int>();
            for (int i = 0; i < points.Count - 1; i++)
            {
                int current = points[i].PointNumber;
                int next = points[i + 1].PointNumber;
                if (next != current + 1)
                {
                    for (int missing = current + 1; missing < next; missing++)
                    {
                        gaps.Add(missing);
                    }
                }
            }

            if (gaps.Any())
            {
                logBuilder.AppendLine("  GAPS:");
                foreach (var gap in gaps.Take(20)) // Show first 20 gaps
                {
                    logBuilder.AppendLine($"    Missing: {gap}");
                }
                if (gaps.Count > 20)
                {
                    logBuilder.AppendLine($"    ... and {gaps.Count - 20} more");
                }
            }

            if (!duplicates.Any() && !gaps.Any())
            {
                logBuilder.AppendLine("  ✓ Continuous sequence 1..N");
            }
        }

        /// <summary>
        /// Write Excel output
        /// </summary>
        private static void WriteExcelOutput(List<PointData> points, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Points");
                
                // Headers
                sheet.Cell(1, 1).Value = "Point Number";
                sheet.Cell(1, 1).Style.Font.Bold = true;
                sheet.Cell(1, 2).Value = "Point Name";
                sheet.Cell(1, 2).Style.Font.Bold = true;

                // Data
                int row = 2;
                foreach (var point in points)
                {
                    sheet.Cell(row, 1).Value = point.PointNumber;
                    sheet.Cell(row, 2).Value = point.PointName;
                    row++;
                }

                workbook.SaveAs(outputPath);
            }
        }
    }

    /// <summary>
    /// Word with bounding box
    /// </summary>
    public class Word
    {
        public string Text { get; set; } = "";
        public int X0 { get; set; }
        public int Y0 { get; set; }
        public int X1 { get; set; }
        public int Y1 { get; set; }
    }

    /// <summary>
    /// Table header definition
    /// </summary>
    public class TableHeader
    {
        public int PointNumberX0 { get; set; }
        public int PointNumberX1 { get; set; }
        public int PointNameX0 { get; set; }
        public int PointNameX1 { get; set; }
        public int HeaderY1 { get; set; }
    }

    /// <summary>
    /// Point data
    /// </summary>
    public class PointData
    {
        public int PointNumber { get; set; }
        public string PointName { get; set; } = "";
    }
}
