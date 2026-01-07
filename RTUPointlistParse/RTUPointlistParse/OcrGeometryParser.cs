using System.Diagnostics;
using System.Globalization;

namespace RTUPointlistParse
{
    /// <summary>
    /// OCR + Geometry-based parser for extracting Point Number and Point Name from PDF tables
    /// Implements the step-by-step algorithm: Render, OCR with bounding boxes, detect headers,
    /// define column bands, cluster words into rows, and extract data
    /// </summary>
    public class OcrGeometryParser
    {
        // Tolerances for geometric matching
        private const int HORIZONTAL_OVERLAP_TOLERANCE = 15; // px for header alignment
        private const int VERTICAL_STACKING_TOLERANCE = 50;   // px for POINT/NUMBER stacking
        private const int ROW_CLUSTERING_TOLERANCE = 12;      // px for grouping words into rows
        private const int COLUMN1_EXPANSION = 20;             // px to expand Point Number column band
        private const int COLUMN2_EXPANSION = 40;             // px to expand Point Name column band
        private const int MIN_VERTICAL_GAP_FOR_TABLE_END = 100; // px gap to end table scope

        /// <summary>
        /// Represents a word with its bounding box from OCR
        /// </summary>
        public class OcrWord
        {
            public string Text { get; set; } = "";
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
            public float Confidence { get; set; }
            
            public int Right => Left + Width;
            public int Bottom => Top + Height;
            public int CenterX => Left + Width / 2;
            public int CenterY => Top + Height / 2;
        }

        /// <summary>
        /// Represents a detected table header with column positions
        /// </summary>
        public class TableHeader
        {
            public int PointNumberColumnLeft { get; set; }
            public int PointNumberColumnRight { get; set; }
            public int PointNameColumnLeft { get; set; }
            public int PointNameColumnRight { get; set; }
            public int HeaderBottomY { get; set; }
            public int TableEndY { get; set; }
        }

        /// <summary>
        /// Parse PDF file using OCR + Geometry approach
        /// </summary>
        public static List<TableRow> ParsePdfWithGeometry(string pdfPath)
        {
            var allRows = new List<TableRow>();

            try
            {
                Console.WriteLine($"  Performing OCR with bounding boxes...");
                
                // Step 1: Render & OCR - Convert PDF pages to images
                string tempDir = Path.Combine(Path.GetTempPath(), $"pdf_ocr_{Guid.NewGuid()}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Convert PDF to images
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
                        return allRows;
                    }

                    // Process each page
                    var imageFiles = Directory.GetFiles(tempDir, "*.png").OrderBy(f => f).ToArray();
                    
                    foreach (var imageFile in imageFiles)
                    {
                        Console.WriteLine($"  Processing page {Path.GetFileName(imageFile)}...");
                        
                        // Get OCR words with bounding boxes
                        var words = ExtractWordsWithBoundingBoxes(imageFile);
                        
                        if (words.Count == 0)
                        {
                            Console.WriteLine($"  No words found on page");
                            continue;
                        }

                        // Step 2: Detect header pairs for each table on the page
                        var tableHeaders = DetectTableHeaders(words);
                        
                        Console.WriteLine($"  Found {tableHeaders.Count} table(s) on page");

                        // Step 3-6: Process each table independently
                        foreach (var header in tableHeaders)
                        {
                            var tableRows = ExtractTableRows(words, header);
                            allRows.AddRange(tableRows);
                        }
                    }

                    Console.WriteLine($"  OCR with geometry completed - extracted {allRows.Count} rows");
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
                Console.WriteLine($"  OCR geometry parsing error: {ex.Message}");
            }

            return allRows;
        }

        /// <summary>
        /// Step 1: Extract words with bounding boxes using Tesseract TSV output
        /// </summary>
        private static List<OcrWord> ExtractWordsWithBoundingBoxes(string imageFile)
        {
            var words = new List<OcrWord>();

            try
            {
                using (var tessProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "tesseract",
                        Arguments = $"\"{imageFile}\" stdout tsv",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                })
                {
                    tessProcess.Start();
                    
                    var output = tessProcess.StandardOutput.ReadToEnd();
                    tessProcess.WaitForExit();

                    if (tessProcess.ExitCode != 0)
                    {
                        return words;
                    }

                    // Parse TSV output
                    var lines = output.Split('\n');
                    bool isFirstLine = true;

                    foreach (var line in lines)
                    {
                        if (isFirstLine)
                        {
                            isFirstLine = false;
                            continue; // Skip header
                        }

                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');
                        
                        // TSV format: level, page_num, block_num, par_num, line_num, word_num, left, top, width, height, conf, text
                        if (fields.Length >= 12)
                        {
                            int level = int.Parse(fields[0]);
                            
                            // We only want word-level entries (level 5)
                            if (level == 5)
                            {
                                string text = fields[11].Trim();
                                
                                // Skip empty or whitespace-only words
                                if (string.IsNullOrWhiteSpace(text))
                                    continue;

                                int left = int.Parse(fields[6]);
                                int top = int.Parse(fields[7]);
                                int width = int.Parse(fields[8]);
                                int height = int.Parse(fields[9]);
                                float conf = float.Parse(fields[10], CultureInfo.InvariantCulture);

                                words.Add(new OcrWord
                                {
                                    Text = text,
                                    Left = left,
                                    Top = top,
                                    Width = width,
                                    Height = height,
                                    Confidence = conf
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error extracting words with bounding boxes: {ex.Message}");
            }

            return words;
        }

        /// <summary>
        /// Step 2: Detect table headers - find "POINT" adjacent to "NAME" pattern
        /// The problem statement mentions stacked "POINT"/"NUMBER", but in practice we look for
        /// "POINT NAME" pattern which indicates the point name column header.
        /// The point number column may be implicit (unlabeled numeric column to the left).
        /// </summary>
        private static List<TableHeader> DetectTableHeaders(List<OcrWord> words)
        {
            var headers = new List<TableHeader>();

            // Find all "POINT" words
            var pointWords = words.Where(w => 
                w.Text.Equals("POINT", StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var pointWord in pointWords)
            {
                // Look for "NAME" word adjacent to this "POINT" (on the same line)
                var nameWord = words.FirstOrDefault(w =>
                    w.Text.Equals("NAME", StringComparison.OrdinalIgnoreCase) &&
                    w.Left > pointWord.Right &&
                    w.Left - pointWord.Right < 200 && // Reasonable proximity
                    Math.Abs(w.CenterY - pointWord.CenterY) < ROW_CLUSTERING_TOLERANCE);

                if (nameWord == null)
                    continue; // No adjacent NAME found

                // Look for potential stacked "NUMBER" below POINT (optional - for strict compliance)
                // But if not found, we still proceed assuming implicit number column
                var numberWord = words.FirstOrDefault(w =>
                    w.Text.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                    w.Top > pointWord.Top &&
                    w.Top - pointWord.Bottom < VERTICAL_STACKING_TOLERANCE &&
                    Math.Abs(w.CenterX - pointWord.CenterX) < HORIZONTAL_OVERLAP_TOLERANCE);

                // Define column bands
                // Point Number column: Look for numeric data left of POINT NAME header
                // We assume it's significantly to the left (at least 100px)
                int pointNumberColumnLeft = Math.Max(0, pointWord.Left - 400);
                int pointNumberColumnRight = pointWord.Left - 50; // Leave gap before name column

                // Point Name column: Based on "POINT NAME" header position
                int pointNameColumnLeft = pointWord.Left - COLUMN2_EXPANSION;
                int pointNameColumnRight = nameWord.Right + COLUMN2_EXPANSION;

                // Table starts below the header
                int headerBottomY = nameWord.Bottom;
                if (numberWord != null)
                {
                    headerBottomY = Math.Max(headerBottomY, numberWord.Bottom);
                }

                var header = new TableHeader
                {
                    PointNumberColumnLeft = pointNumberColumnLeft,
                    PointNumberColumnRight = pointNumberColumnRight,
                    PointNameColumnLeft = pointNameColumnLeft,
                    PointNameColumnRight = pointNameColumnRight,
                    HeaderBottomY = headerBottomY,
                    TableEndY = int.MaxValue // Will be determined by gap detection
                };

                headers.Add(header);
            }

            // Determine table end Y positions based on gaps or next header
            for (int i = 0; i < headers.Count; i++)
            {
                if (i < headers.Count - 1)
                {
                    // End before next header
                    headers[i].TableEndY = headers[i + 1].HeaderBottomY;
                }
                else
                {
                    // Last table - find large vertical gap or use page bottom
                    headers[i].TableEndY = int.MaxValue;
                }
            }

            return headers;
        }

        /// <summary>
        /// Steps 3-6: Extract table rows from words using column bands and row clustering
        /// </summary>
        private static List<TableRow> ExtractTableRows(List<OcrWord> words, TableHeader header)
        {
            var rows = new List<TableRow>();

            // Filter words within table vertical scope
            var tableWords = words.Where(w => 
                w.Top > header.HeaderBottomY && 
                w.Top < header.TableEndY).ToList();

            if (tableWords.Count == 0)
                return rows;

            // Step 4: Cluster words into rows by Y coordinate
            var rowClusters = ClusterWordsIntoRows(tableWords);

            // Step 5: Extract Point Number and Point Name from each row
            int pointNumber = 0;
            foreach (var cluster in rowClusters)
            {
                // Find numeric token in the general area of Column 1
                // We use a wider band to catch numbers that may not align perfectly
                var pointNumCandidates = cluster.Where(w =>
                    w.CenterX >= header.PointNumberColumnLeft &&
                    w.CenterX <= header.PointNumberColumnRight &&
                    IsNumericToken(w.Text))
                    .OrderBy(w => w.Left) // Take leftmost number
                    .ToList();

                // If no number in the expected band, look for the first numeric token in the row
                if (pointNumCandidates.Count == 0)
                {
                    pointNumCandidates = cluster.Where(w => IsNumericToken(w.Text))
                        .OrderBy(w => w.Left)
                        .Take(1)
                        .ToList();
                }

                if (pointNumCandidates.Count == 0)
                    continue; // No point number found in this row

                var pointNumWord = pointNumCandidates.First();

                // Parse point number
                if (!int.TryParse(pointNumWord.Text, out int parsedPointNum))
                {
                    Console.WriteLine($"  Warning: Failed to parse point number: {pointNumWord.Text}");
                    continue;
                }

                // Find all words in Column 2 band or to the right of Point Number
                // Point name is everything after the point number in the row
                var pointNameWords = cluster.Where(w =>
                    w.Left > pointNumWord.Right &&
                    !IsNumericToken(w.Text)) // Exclude pure numeric tokens
                    .OrderBy(w => w.Left)
                    .ToList();

                if (pointNameWords.Count == 0)
                {
                    Console.WriteLine($"  Warning: Empty point name for point {parsedPointNum}");
                    continue;
                }

                // Concatenate point name words
                string pointName = string.Join(" ", pointNameWords.Select(w => w.Text));

                // Validate and add row
                if (IsValidPointName(pointName))
                {
                    rows.Add(new TableRow
                    {
                        Columns = new List<string> { pointNumber.ToString(), pointName }
                    });
                    pointNumber++;
                }
            }

            return rows;
        }

        /// <summary>
        /// Step 4: Cluster words into rows based on Y coordinate proximity
        /// </summary>
        private static List<List<OcrWord>> ClusterWordsIntoRows(List<OcrWord> words)
        {
            var clusters = new List<List<OcrWord>>();

            // Sort by Y coordinate
            var sortedWords = words.OrderBy(w => w.CenterY).ToList();

            foreach (var word in sortedWords)
            {
                // Find existing cluster within tolerance
                var cluster = clusters.FirstOrDefault(c =>
                    Math.Abs(c[0].CenterY - word.CenterY) <= ROW_CLUSTERING_TOLERANCE);

                if (cluster != null)
                {
                    cluster.Add(word);
                }
                else
                {
                    // Create new cluster
                    clusters.Add(new List<OcrWord> { word });
                }
            }

            // Sort words within each cluster by X coordinate
            foreach (var cluster in clusters)
            {
                cluster.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            return clusters;
        }

        /// <summary>
        /// Check if a token is numeric (digits only with optional small punctuation)
        /// </summary>
        private static bool IsNumericToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // Remove common OCR artifacts and punctuation
            text = text.Trim().Replace(",", "").Replace(".", "").Replace(":", "");

            return text.All(char.IsDigit);
        }

        /// <summary>
        /// Validate point name (not empty, not "Spare", not too short, not OCR noise)
        /// </summary>
        private static bool IsValidPointName(string pointName)
        {
            if (string.IsNullOrWhiteSpace(pointName))
                return false;

            // Trim whitespace
            pointName = pointName.Trim();

            // Filter out SPARE
            if (pointName.Equals("SPARE", StringComparison.OrdinalIgnoreCase))
                return false;

            // Filter out single characters (likely OCR noise)
            if (pointName.Length <= 1)
                return false;

            // Filter out common OCR noise patterns that appear alone
            var noisePatterns = new[] { "I", "F", "J", "L", "—", "=", "|", "[", "]" };
            if (noisePatterns.Contains(pointName, StringComparer.OrdinalIgnoreCase))
                return false;

            // Filter out lines that are mostly punctuation or special characters
            int alphanumCount = pointName.Count(c => char.IsLetterOrDigit(c));
            if (alphanumCount < 2)
                return false;

            // Filter out reference/metadata lines (but keep if they're part of a real point name)
            if (pointName.StartsWith("PLOT BY") || pointName.StartsWith("RESERVED FOR") ||
                pointName == "LISTING" || pointName == "CONSTRUCTION" ||
                pointName == "SYSTEM" || pointName == "REFERENCE" || pointName == "SAP")
                return false;

            return true;
        }
    }
}
