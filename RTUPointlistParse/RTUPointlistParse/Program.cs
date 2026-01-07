using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using ClosedXML.Excel;
using Docnet.Core;
using Docnet.Core.Models;
using Tesseract;

namespace RTUPointlistParse
{
    public record OcrWord(string Text, Rectangle Bounds, float Confidence);
    public record TableHeader(Rectangle Bounds, int PageIndex, string SourceFile);
    public record RowCluster(int PageIndex, string SourceFile, int Y, List<OcrWord> Words);

    public class App
    {
        public const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
        public const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";
        private const int DPI = 300;
        private const int Y_TOLERANCE = 10;
        private const int HEADER_X_TOLERANCE = 20;
        private const int COLUMN_X_TOLERANCE = 15;
        private const int MAX_VERTICAL_GAP = 50;

        private readonly List<string> logLines = new();

        public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVar, string? tessLangEnvVar)
        {
            try
            {
                Log("RTU Point List Parser - OCR-Based");
                Log("=====================================");
                Log($"Input folder: {inputFolder}");
                Log($"Output folder: {outputFolder}");
                Log("");

                // Validate input folder
                if (!Directory.Exists(inputFolder))
                {
                    Log($"ERROR: Input folder does not exist: {inputFolder}");
                    return 1;
                }

                // Create output folder
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                    Log($"Created output folder: {outputFolder}");
                }

                // Get tessdata path and language
                string? tessDataDir = Environment.GetEnvironmentVariable(tessDataEnvVar ?? "TESSDATA_DIR");
                string tessLang = Environment.GetEnvironmentVariable(tessLangEnvVar ?? "TESSERACT_LANG") ?? "eng";

                if (string.IsNullOrWhiteSpace(tessDataDir))
                {
                    Log("ERROR: TESSDATA_DIR environment variable not set.");
                    Log("Please set it to the path containing tessdata folder (e.g., C:\\tools\\tesseract\\tessdata)");
                    return 1;
                }

                if (!Directory.Exists(tessDataDir))
                {
                    Log($"ERROR: Tessdata directory not found: {tessDataDir}");
                    return 1;
                }

                string trainedDataPath = Path.Combine(tessDataDir, $"{tessLang}.traineddata");
                if (!File.Exists(trainedDataPath))
                {
                    Log($"ERROR: Trained data file not found: {trainedDataPath}");
                    Log($"Please ensure {tessLang}.traineddata exists in the tessdata folder.");
                    return 1;
                }

                Log($"Tesseract data: {tessDataDir}");
                Log($"Language: {tessLang}");
                Log("");

                // Get all PDF files
                var pdfFiles = GetPdfFiles(inputFolder);
                Log($"Found {pdfFiles.Count} PDF file(s)");

                if (pdfFiles.Count == 0)
                {
                    Log("WARNING: No PDF files found in input folder");
                    WriteLog(outputFolder, logLines);
                    return 0;
                }

                // Collect all extracted rows
                var allRows = new List<(int PointNumber, string PointName)>();

                // Initialize Tesseract engine (reuse for all pages)
                using var engine = new TesseractEngine(tessDataDir, tessLang, EngineMode.Default);
                engine.SetVariable("tessedit_pageseg_mode", "6"); // Assume uniform block of text

                // Process each PDF
                foreach (var pdfPath in pdfFiles)
                {
                    var fileName = Path.GetFileName(pdfPath);
                    Log($"Processing: {fileName}");

                    try
                    {
                        var startTime = DateTime.Now;
                        
                        // Render PDF to bitmaps
                        var bitmaps = RenderPdfToBitmaps(pdfPath, DPI, Log);
                        Log($"  Rendered {bitmaps.Count} page(s) to bitmaps");

                        // Process each page
                        for (int pageIdx = 0; pageIdx < bitmaps.Count; pageIdx++)
                        {
                            using var bitmap = bitmaps[pageIdx];
                            var pageStartTime = DateTime.Now;

                            // Run OCR
                            var words = OcrWordsFromBitmap(bitmap, engine);
                            var elapsed = (DateTime.Now - pageStartTime).TotalSeconds;
                            Log($"  Page {pageIdx + 1}: OCR extracted {words.Count} words in {elapsed:F2}s");

                            if (words.Count == 0)
                            {
                                Log($"  Page {pageIdx + 1}: WARNING: No words extracted");
                                continue;
                            }

                            // Detect table headers
                            var headers = DetectPointNumberHeaders(words);
                            Log($"  Page {pageIdx + 1}: Detected {headers.Count} 'Point Number' header(s)");

                            if (headers.Count == 0)
                            {
                                Log($"  Page {pageIdx + 1}: WARNING: No table headers found");
                                continue;
                            }

                            // Extract rows from each header
                            foreach (var header in headers)
                            {
                                var headerCopy = new TableHeader(header.Bounds, pageIdx, fileName);
                                var rows = ExtractRowsFromHeaders(words, new List<TableHeader> { headerCopy }, Log);
                                allRows.AddRange(rows);
                            }
                        }

                        var totalElapsed = (DateTime.Now - startTime).TotalSeconds;
                        Log($"  Total processing time: {totalElapsed:F2}s");
                        Log($"  Extracted {allRows.Count} total rows so far");
                    }
                    catch (Exception ex)
                    {
                        Log($"  ERROR processing {fileName}: {ex.Message}");
                    }
                }

                Log("");
                Log($"Total rows extracted from all PDFs: {allRows.Count}");

                if (allRows.Count == 0)
                {
                    Log("WARNING: No data rows extracted");
                    WriteLog(outputFolder, logLines);
                    return 0;
                }

                // Sort by Point Number
                var sortedRows = allRows.OrderBy(r => r.PointNumber).ToList();
                Log($"Sorted {sortedRows.Count} rows by Point Number");

                // Validate sequence
                var pointNumbers = sortedRows.Select(r => r.PointNumber).ToList();
                var (duplicates, gaps) = ValidateSequence(pointNumbers);

                Log("");
                Log("Sequence Validation:");
                Log($"  Expected range: 1 to {pointNumbers.Max()}");
                Log($"  Duplicates found: {duplicates.Count}");
                Log($"  Gaps found: {gaps.Count}");

                if (duplicates.Count > 0)
                {
                    Log($"  First 50 duplicates: {string.Join(", ", duplicates.Take(50))}");
                }

                if (gaps.Count > 0)
                {
                    Log($"  First 50 missing numbers: {string.Join(", ", gaps.Take(50))}");
                }

                // Write Excel output
                string outputXlsxPath = Path.Combine(outputFolder, "PointList.xlsx");
                WriteExcel(outputXlsxPath, sortedRows);
                Log("");
                Log($"Excel file created: {outputXlsxPath}");

                // Write log file
                WriteLog(outputFolder, logLines);
                Log($"Log file created: {Path.Combine(outputFolder, "PointList.log.txt")}");

                Log("");
                Log("Processing complete.");
                return 0;
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                WriteLog(outputFolder, logLines);
                return 1;
            }
        }

        private void Log(string message)
        {
            Console.WriteLine(message);
            logLines.Add(message);
        }

        public static List<string> GetPdfFiles(string inputFolder)
        {
            return Directory.GetFiles(inputFolder, "*.pdf", SearchOption.TopDirectoryOnly).ToList();
        }

        public static List<Bitmap> RenderPdfToBitmaps(string pdfPath, int dpi, Action<string> log)
        {
            var bitmaps = new List<Bitmap>();

            try
            {
                using var library = DocLib.Instance;
                using var docReader = library.GetDocReader(pdfPath, new PageDimensions(dpi, dpi));
                
                for (int i = 0; i < docReader.GetPageCount(); i++)
                {
                    using var pageReader = docReader.GetPageReader(i);
                    
                    // Get page dimensions
                    int width = pageReader.GetPageWidth();
                    int height = pageReader.GetPageHeight();
                    
                    // Render page as raw bytes (BGRA format)
                    byte[] bytes = pageReader.GetImage();
                    
                    // Create bitmap from raw bytes
                    var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                    var bitmapData = bitmap.LockBits(
                        new Rectangle(0, 0, width, height),
                        ImageLockMode.WriteOnly,
                        PixelFormat.Format32bppArgb);
                    
                    System.Runtime.InteropServices.Marshal.Copy(bytes, 0, bitmapData.Scan0, bytes.Length);
                    bitmap.UnlockBits(bitmapData);
                    
                    bitmaps.Add(bitmap);
                }
            }
            catch (Exception ex)
            {
                log($"  ERROR rendering PDF: {ex.Message}");
            }

            return bitmaps;
        }

        public static List<OcrWord> OcrWordsFromBitmap(Bitmap bmp, TesseractEngine engine)
        {
            var words = new List<OcrWord>();

            try
            {
                // Convert Bitmap to Pix
                using var pix = BitmapToPix(bmp);
                using var page = engine.Process(pix);

                using var iterator = page.GetIterator();
                iterator.Begin();

                do
                {
                    if (iterator.TryGetBoundingBox(PageIteratorLevel.Word, out var bounds))
                    {
                        string text = iterator.GetText(PageIteratorLevel.Word);
                        float confidence = iterator.GetConfidence(PageIteratorLevel.Word);

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var rect = new Rectangle(bounds.X1, bounds.Y1, bounds.Width, bounds.Height);
                            words.Add(new OcrWord(text.Trim(), rect, confidence));
                        }
                    }
                } while (iterator.Next(PageIteratorLevel.Word));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR during OCR: {ex.Message}");
            }

            return words;
        }

        private static Pix BitmapToPix(Bitmap bitmap)
        {
            // Save bitmap to memory stream and load as Pix
            using var ms = new MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;
            return Pix.LoadFromMemory(ms.ToArray());
        }

        public static List<TableHeader> DetectPointNumberHeaders(List<OcrWord> words)
        {
            var headers = new List<TableHeader>();

            // Find all "Point" words
            var pointWords = words.Where(w => 
                string.Equals(w.Text, "Point", StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var pointWord in pointWords)
            {
                // Find "Number" words below and overlapping horizontally
                var numberWords = words.Where(w =>
                    string.Equals(w.Text, "Number", StringComparison.OrdinalIgnoreCase) &&
                    w.Bounds.Top > pointWord.Bounds.Top &&
                    w.Bounds.Top < pointWord.Bounds.Top + 100 && // Within reasonable vertical distance
                    Math.Abs(w.Bounds.Left - pointWord.Bounds.Left) < HEADER_X_TOLERANCE // Horizontally aligned
                ).ToList();

                foreach (var numberWord in numberWords)
                {
                    // Create combined bounding box
                    int left = Math.Min(pointWord.Bounds.Left, numberWord.Bounds.Left);
                    int top = pointWord.Bounds.Top;
                    int right = Math.Max(pointWord.Bounds.Right, numberWord.Bounds.Right);
                    int bottom = numberWord.Bounds.Bottom;
                    
                    var combinedBounds = new Rectangle(left, top, right - left, bottom - top);
                    headers.Add(new TableHeader(combinedBounds, 0, ""));
                }
            }

            return headers;
        }

        public static IEnumerable<(int PointNumber, string PointName)> ExtractRowsFromHeaders(
            List<OcrWord> words, List<TableHeader> headers, Action<string> log)
        {
            var result = new List<(int PointNumber, string PointName)>();

            foreach (var header in headers)
            {
                // Cluster words into rows
                var rowClusters = ClusterWordsIntoRows(words, Y_TOLERANCE);

                // Filter rows below the header
                var dataRows = rowClusters
                    .Where(rc => rc.Y > header.Bounds.Bottom)
                    .OrderBy(rc => rc.Y)
                    .ToList();

                // Process each row
                int lastY = header.Bounds.Bottom;
                foreach (var row in dataRows)
                {
                    // Stop if there's a large vertical gap (indicates end of table)
                    if (row.Y - lastY > MAX_VERTICAL_GAP)
                        break;

                    lastY = row.Y;

                    // Find Point Number in this row (numeric word within header x-range)
                    var pointNumberWord = row.Words
                        .Where(w => int.TryParse(w.Text, out _))
                        .Where(w => Math.Abs(w.Bounds.Left - header.Bounds.Left) < COLUMN_X_TOLERANCE * 2)
                        .FirstOrDefault();

                    if (pointNumberWord == null)
                        continue;

                    if (!int.TryParse(pointNumberWord.Text, out int pointNumber))
                        continue;

                    // Find Point Name (all words to the right of Point Number)
                    var nameWords = row.Words
                        .Where(w => w.Bounds.Left > pointNumberWord.Bounds.Right)
                        .OrderBy(w => w.Bounds.Left)
                        .ToList();

                    if (nameWords.Count == 0)
                        continue;

                    string pointName = NormalizeText(string.Join(" ", nameWords.Select(w => w.Text)));

                    if (string.IsNullOrWhiteSpace(pointName) || pointName.Length <= 2)
                        continue;

                    result.Add((pointNumber, pointName));
                }
            }

            return result;
        }

        public static IEnumerable<RowCluster> ClusterWordsIntoRows(List<OcrWord> words, int yTolerance)
        {
            var clusters = new List<RowCluster>();

            // Group words by approximate Y coordinate (center)
            var grouped = words
                .GroupBy(w => (w.Bounds.Top + w.Bounds.Height / 2) / yTolerance)
                .OrderBy(g => g.Key);

            foreach (var group in grouped)
            {
                int avgY = (int)group.Average(w => w.Bounds.Top + w.Bounds.Height / 2);
                var wordsInRow = group.ToList();
                clusters.Add(new RowCluster(0, "", avgY, wordsInRow));
            }

            return clusters;
        }

        public static string NormalizeText(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";

            // Trim and collapse multiple spaces
            s = s.Trim();
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ");
            return s;
        }

        public static (List<int> duplicates, List<int> gaps) ValidateSequence(IEnumerable<int> nums)
        {
            var duplicates = new List<int>();
            var gaps = new List<int>();

            var numList = nums.ToList();
            if (numList.Count == 0)
                return (duplicates, gaps);

            // Find duplicates
            var seen = new HashSet<int>();
            foreach (var num in numList)
            {
                if (!seen.Add(num))
                    duplicates.Add(num);
            }

            // Find gaps
            int maxNum = numList.Max();
            var uniqueNums = new HashSet<int>(numList);
            
            for (int i = 1; i <= maxNum; i++)
            {
                if (!uniqueNums.Contains(i))
                    gaps.Add(i);
            }

            return (duplicates.Distinct().ToList(), gaps);
        }

        public static void WriteExcel(string outputXlsxPath, IEnumerable<(int PointNumber, string PointName)> rows)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Points");

            // Header
            worksheet.Cell(1, 1).Value = "Point Number";
            worksheet.Cell(1, 2).Value = "Point Name";
            worksheet.Row(1).Style.Font.Bold = true;

            // Data
            int rowNum = 2;
            foreach (var (pointNumber, pointName) in rows)
            {
                worksheet.Cell(rowNum, 1).Value = pointNumber;
                worksheet.Cell(rowNum, 2).Value = pointName;
                rowNum++;
            }

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(outputXlsxPath);
        }

        public static void WriteLog(string outputFolder, IEnumerable<string> logLines)
        {
            string logPath = Path.Combine(outputFolder, "PointList.log.txt");
            File.WriteAllLines(logPath, logLines);
        }
    }

    public class Program
    {
        public static async Task<int> Main(string[] args)
        {
            var app = new App();
            return await app.RunAsync(
                App.DefaultInputFolder,
                App.DefaultOutputFolder,
                "TESSDATA_DIR",
                "TESSERACT_LANG");
        }
    }
}
