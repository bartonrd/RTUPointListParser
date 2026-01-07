using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using PdfiumViewer;
using Tesseract;
using ClosedXML.Excel;

namespace RTUPointlistParse;

public class App
{
    private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
    private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";
    
    private readonly List<string> _logLines = new();

    public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVar, string? tessLangEnvVar)
    {
        return await Task.Run(() =>
        {
            try
            {
                Log("RTU Point List Parser");
                Log("=====================");
                Log($"Input folder: {inputFolder}");
                Log($"Output folder: {outputFolder}");
                Log("");

                if (!Directory.Exists(inputFolder))
                {
                    Log($"ERROR: Input folder does not exist: {inputFolder}");
                    return 1;
                }

                Directory.CreateDirectory(outputFolder);

                var pdfFiles = GetPdfFiles(inputFolder);
                Log($"Found {pdfFiles.Count} PDF file(s) to process");
                Log("");

                if (pdfFiles.Count == 0)
                {
                    Log("No PDF files found.");
                    return 1;
                }

                var tessDataPath = Environment.GetEnvironmentVariable(tessDataEnvVar ?? "TESSDATA_DIR") ?? 
                                   Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", "tessdata");
                var tessLang = Environment.GetEnvironmentVariable(tessLangEnvVar ?? "TESSERACT_LANG") ?? "eng";

                if (!Directory.Exists(tessDataPath))
                {
                    Log($"ERROR: Tessdata directory not found: {tessDataPath}");
                    Log($"Set {tessDataEnvVar ?? "TESSDATA_DIR"} environment variable or place tessdata in project root.");
                    return 1;
                }

                var allRows = new List<(int PointNumber, string PointName)>();

                using (var engine = new TesseractEngine(tessDataPath, tessLang, EngineMode.Default))
                {
                    foreach (var pdfPath in pdfFiles)
                    {
                        Log($"Processing: {Path.GetFileName(pdfPath)}");

                        var bitmaps = RenderPdfToBitmaps(pdfPath, 300, Log);
                        Log($"  Rendered {bitmaps.Count} page(s)");

                        foreach (var bitmap in bitmaps)
                        {
                            try
                            {
                                var words = OcrWordsFromBitmap(bitmap, engine);
                                var bands = DetectFirstTwoColumnBands(words);

                                if (bands != null)
                                {
                                    var rows = ExtractRows(words, bands.Value, Log);
                                    allRows.AddRange(rows);
                                }
                            }
                            finally
                            {
                                bitmap.Dispose();
                            }
                        }

                        Log($"  Extracted {allRows.Count} total rows so far");
                    }
                }

                Log("");
                Log($"Total rows extracted: {allRows.Count}");

                // Sort by Point Number
                allRows = allRows.OrderBy(r => r.PointNumber).ToList();

                // Validate sequence
                var pointNumbers = allRows.Select(r => r.PointNumber).ToList();
                var (duplicates, gaps) = ValidateSequence(pointNumbers);

                if (duplicates.Count > 0)
                {
                    Log($"WARNING: Found {duplicates.Count} duplicate point numbers: {string.Join(", ", duplicates)}");
                }

                if (gaps.Count > 0)
                {
                    Log($"WARNING: Found {gaps.Count} gaps in sequence: {string.Join(", ", gaps)}");
                }

                // Write Excel output
                var outputXlsxPath = Path.Combine(outputFolder, "PointList.xlsx");
                WriteExcel(outputXlsxPath, allRows);
                Log("");
                Log($"Written: {outputXlsxPath}");

                // Write log
                WriteLog(outputFolder, _logLines);
                Log($"Written: {Path.Combine(outputFolder, "PointList.log.txt")}");

                return 0;
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR: {ex.Message}");
                Log(ex.StackTrace ?? "");
                WriteLog(outputFolder, _logLines);
                return 1;
            }
        });
    }

    private void Log(string message)
    {
        Console.WriteLine(message);
        _logLines.Add(message);
    }

    public static List<string> GetPdfFiles(string inputFolder)
    {
        return Directory.GetFiles(inputFolder, "*.pdf").OrderBy(f => f).ToList();
    }

    public static List<Bitmap> RenderPdfToBitmaps(string pdfPath, int dpi, Action<string> log)
    {
        var bitmaps = new List<Bitmap>();

        try
        {
            using var document = PdfDocument.Load(pdfPath);
            for (int i = 0; i < document.PageCount; i++)
            {
                var image = document.Render(i, dpi, dpi, PdfRenderFlags.Annotations);
                bitmaps.Add(new Bitmap(image));
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

        // Convert Bitmap to byte array for Tesseract
        using (var ms = new System.IO.MemoryStream())
        {
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;
            
            using (var pix = Tesseract.Pix.LoadFromMemory(ms.ToArray()))
            using (var page = engine.Process(pix))
            {
                using var iter = page.GetIterator();
                iter.Begin();

                do
                {
                    if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var bounds))
                    {
                        var text = iter.GetText(PageIteratorLevel.Word);
                        var confidence = iter.GetConfidence(PageIteratorLevel.Word);

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            words.Add(new OcrWord(
                                text.Trim(),
                                new Rectangle(bounds.X1, bounds.Y1, bounds.X2 - bounds.X1, bounds.Y2 - bounds.Y1),
                                confidence
                            ));
                        }
                    }
                } while (iter.Next(PageIteratorLevel.Word));
            }
        }

        return words;
    }

    public static (Rectangle col1, Rectangle col2)? DetectFirstTwoColumnBands(List<OcrWord> words)
    {
        if (words.Count == 0) return null;

        // Find header words containing "POINT" and "NUMBER" (may be stacked)
        var pointWords = words.Where(w => 
            w.Text.Contains("POINT", StringComparison.OrdinalIgnoreCase) ||
            w.Text.Contains("NUMBER", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        if (pointWords.Count == 0)
        {
            // Fallback: assume leftmost column
            var minX = words.Min(w => w.Bounds.X);
            var maxX = words.Max(w => w.Bounds.Right);
            var minY = words.Min(w => w.Bounds.Y);
            var maxY = words.Max(w => w.Bounds.Bottom);

            var midX = minX + (maxX - minX) / 3;

            return (
                new Rectangle(minX, minY, midX - minX, maxY - minY),
                new Rectangle(midX, minY, maxX - midX, maxY - minY)
            );
        }

        // Find the leftmost header column (Point Number)
        var col1X = pointWords.Min(w => w.Bounds.X);
        var col1Width = 150; // approximate width for point number column

        // Find the next column to the right (Point Name)
        var nameWords = words.Where(w => 
            w.Text.Contains("NAME", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        var col2X = nameWords.Count > 0 
            ? nameWords.Min(w => w.Bounds.X)
            : col1X + col1Width;

        var pageHeight = words.Max(w => w.Bounds.Bottom);
        var col2Width = words.Max(w => w.Bounds.Right) - col2X;

        return (
            new Rectangle(col1X - 10, 0, col1Width, pageHeight),
            new Rectangle(col2X - 10, 0, col2Width, pageHeight)
        );
    }

    public static IEnumerable<RowCluster> ClusterWordsIntoRows(List<OcrWord> words, int yTol)
    {
        var clusters = new Dictionary<int, List<OcrWord>>();

        foreach (var word in words)
        {
            var yKey = (word.Bounds.Y / yTol) * yTol;
            if (!clusters.ContainsKey(yKey))
                clusters[yKey] = new List<OcrWord>();
            clusters[yKey].Add(word);
        }

        return clusters.Select(kvp => new RowCluster(kvp.Key, kvp.Value.OrderBy(w => w.Bounds.X).ToList()));
    }

    public static IEnumerable<(int PointNumber, string PointName)> ExtractRows(
        List<OcrWord> words, 
        (Rectangle col1, Rectangle col2) bands, 
        Action<string> log)
    {
        var rows = new List<(int PointNumber, string PointName)>();

        // Cluster words into rows
        var rowClusters = ClusterWordsIntoRows(words, 15).OrderBy(r => r.Y);

        foreach (var row in rowClusters)
        {
            // Get words in column 1 (Point Number)
            var col1Words = row.Words.Where(w => bands.col1.IntersectsWith(w.Bounds)).ToList();
            // Get words in column 2 (Point Name)
            var col2Words = row.Words.Where(w => bands.col2.IntersectsWith(w.Bounds)).ToList();

            if (col1Words.Count == 0) continue;

            // Find first numeric token in column 1
            var pointNumberText = col1Words
                .Select(w => Normalize(w.Text))
                .FirstOrDefault(t => int.TryParse(t, out _));

            if (pointNumberText == null) continue;
            if (!int.TryParse(pointNumberText, out var pointNumber)) continue;

            // Combine all words in column 2 for Point Name
            var pointName = string.Join(" ", col2Words.Select(w => Normalize(w.Text)));

            if (string.IsNullOrWhiteSpace(pointName)) continue;

            // Skip "SPARE" entries
            if (pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase)) continue;

            rows.Add((pointNumber, pointName));
        }

        return rows;
    }

    public static string Normalize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        
        // Remove common OCR artifacts
        s = Regex.Replace(s, @"[|\[\]_]", " ");
        s = Regex.Replace(s, @"\s+", " ");
        
        return s.Trim();
    }

    public static (List<int> duplicates, List<int> gaps) ValidateSequence(IEnumerable<int> nums)
    {
        var numList = nums.ToList();
        var duplicates = new List<int>();
        var gaps = new List<int>();

        if (numList.Count == 0) return (duplicates, gaps);

        // Find duplicates
        var seen = new HashSet<int>();
        foreach (var num in numList)
        {
            if (!seen.Add(num))
                duplicates.Add(num);
        }

        // Find gaps (assuming sequence should be 1..N)
        var distinct = numList.Distinct().OrderBy(n => n).ToList();
        if (distinct.Count > 0)
        {
            for (int i = distinct[0]; i < distinct[^1]; i++)
            {
                if (!distinct.Contains(i))
                    gaps.Add(i);
            }
        }

        return (duplicates.Distinct().ToList(), gaps);
    }

    public static void WriteExcel(string outputXlsxPath, IEnumerable<(int PointNumber, string PointName)> rows)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Points");

        // Headers
        worksheet.Cell(1, 1).Value = "Point Number";
        worksheet.Cell(1, 2).Value = "Point Name";
        worksheet.Cell(1, 1).Style.Font.Bold = true;
        worksheet.Cell(1, 2).Style.Font.Bold = true;

        // Data
        int currentRow = 2;
        foreach (var (pointNumber, pointName) in rows)
        {
            worksheet.Cell(currentRow, 1).Value = pointNumber;
            worksheet.Cell(currentRow, 2).Value = pointName;
            currentRow++;
        }

        workbook.SaveAs(outputXlsxPath);
    }

    public static void WriteLog(string outputFolder, IEnumerable<string> logLines)
    {
        var logPath = Path.Combine(outputFolder, "PointList.log.txt");
        File.WriteAllLines(logPath, logLines);
    }

    public static async Task Main(string[] args)
    {
        var inputFolder = args.Length > 0 ? args[0] : DefaultInputFolder;
        var outputFolder = args.Length > 1 ? args[1] : DefaultOutputFolder;

        var app = new App();
        var exitCode = await app.RunAsync(inputFolder, outputFolder, "TESSDATA_DIR", "TESSERACT_LANG");
        Environment.Exit(exitCode);
    }
}

public record OcrWord(string Text, Rectangle Bounds, float Confidence);

public record RowCluster(int Y, List<OcrWord> Words);
