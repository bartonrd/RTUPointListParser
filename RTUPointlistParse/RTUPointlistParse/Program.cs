using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using PdfiumViewer;
using Tesseract;
using ClosedXML.Excel;

namespace RTUPointlistParse;

public record OcrWord(string Text, Rectangle Bounds, float Confidence);

public class RowCluster
{
    public int Y { get; set; }
    public List<OcrWord> Words { get; set; } = new();
    
    public RowCluster(int y, List<OcrWord> words)
    {
        Y = y;
        Words = words;
    }
}

public class App
{
    private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
    private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";
    private const string TessDataEnvVarName = "TESSDATA_DIR";
    private const string TessLangEnvVarName = "TESSERACT_LANG";

    private readonly List<string> logLines = new();

    public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVar, string? tessLangEnvVar)
    {
        await Task.CompletedTask; // Make async for future extensibility

        Log("RTU Point List Parser");
        Log("=====================");
        Log($"Input folder: {inputFolder}");
        Log($"Output folder: {outputFolder}");
        Log("");

        if (!Directory.Exists(inputFolder))
        {
            Log($"Error: Input folder does not exist: {inputFolder}");
            WriteLog(outputFolder, logLines);
            return 1;
        }

        Directory.CreateDirectory(outputFolder);

        var pdfFiles = GetPdfFiles(inputFolder);
        Log($"Found {pdfFiles.Count} PDF file(s)");
        Log("");

        if (pdfFiles.Count == 0)
        {
            Log("No PDF files found.");
            WriteLog(outputFolder, logLines);
            return 0;
        }

        var allRows = new List<(int PointNumber, string PointName)>();

        var tessDataPath = Environment.GetEnvironmentVariable(tessDataEnvVar ?? "TESSDATA_DIR");
        var tessLang = Environment.GetEnvironmentVariable(tessLangEnvVar ?? "TESSERACT_LANG") ?? "eng";

        if (string.IsNullOrEmpty(tessDataPath))
        {
            Log("Warning: TESSDATA_DIR not set. Tesseract may fail. Please set environment variable to tessdata folder.");
            tessDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        }

        if (!Directory.Exists(tessDataPath))
        {
            Log($"Warning: tessdata path does not exist: {tessDataPath}");
        }

        using var engine = new TesseractEngine(tessDataPath, tessLang, EngineMode.Default);

        foreach (var pdfPath in pdfFiles)
        {
            Log($"Processing: {Path.GetFileName(pdfPath)}");
            try
            {
                var bitmaps = RenderPdfToBitmaps(pdfPath, 300, Log);
                Log($"  Rendered {bitmaps.Count} page(s)");

                foreach (var (bitmap, pageIdx) in bitmaps.Select((b, i) => (b, i)))
                {
                    try
                    {
                        var words = OcrWordsFromBitmap(bitmap, engine);
                        Log($"  Page {pageIdx + 1}: {words.Count} words detected");

                        var bands = DetectFirstTwoColumnBands(words);
                        if (bands == null)
                        {
                            Log($"  Page {pageIdx + 1}: Could not detect two-column layout");
                            continue;
                        }

                        var rows = ExtractRows(words, bands.Value, Log);
                        allRows.AddRange(rows);
                        Log($"  Page {pageIdx + 1}: Extracted {rows.Count()} rows");
                    }
                    finally
                    {
                        bitmap.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  Error: {ex.Message}");
            }
        }

        Log("");
        Log($"Total rows extracted: {allRows.Count}");

        // Sort by Point Number
        var sorted = allRows.OrderBy(r => r.PointNumber).ToList();

        // Validate sequence
        var pointNumbers = sorted.Select(r => r.PointNumber).ToList();
        var (duplicates, gaps) = ValidateSequence(pointNumbers);

        if (duplicates.Count > 0)
        {
            Log($"Warning: Found {duplicates.Count} duplicate point numbers: {string.Join(", ", duplicates.Take(10))}");
        }

        if (gaps.Count > 0)
        {
            Log($"Warning: Found {gaps.Count} gaps in sequence: {string.Join(", ", gaps.Take(10))}");
        }

        // Write Excel
        var xlsxPath = Path.Combine(outputFolder, "PointList.xlsx");
        WriteExcel(xlsxPath, sorted);
        Log($"Excel output: {xlsxPath}");

        // Write log
        var logPath = Path.Combine(outputFolder, "PointList.log.txt");
        Log($"Log output: {logPath}");
        WriteLog(outputFolder, logLines);

        return 0;
    }

    private void Log(string message)
    {
        Console.WriteLine(message);
        logLines.Add(message);
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
                using var image = document.Render(i, dpi, dpi, false);
                bitmaps.Add(new Bitmap(image));
            }
        }
        catch (Exception ex)
        {
            log($"  Error rendering PDF: {ex.Message}");
        }
        return bitmaps;
    }

    public static List<OcrWord> OcrWordsFromBitmap(Bitmap bmp, TesseractEngine engine)
    {
        var words = new List<OcrWord>();
        
        // Convert Bitmap to byte array
        using var ms = new System.IO.MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        var imgBytes = ms.ToArray();
        
        using var pix = Pix.LoadFromMemory(imgBytes);
        using var page = engine.Process(pix, PageSegMode.Auto);
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
                    words.Add(new OcrWord(text.Trim(), new Rectangle(bounds.X1, bounds.Y1, bounds.Width, bounds.Height), confidence));
                }
            }
        } while (iter.Next(PageIteratorLevel.Word));

        return words;
    }

    public static (Rectangle col1, Rectangle col2)? DetectFirstTwoColumnBands(List<OcrWord> words)
    {
        if (words.Count < 10) return null;

        // Group words by X position to find vertical columns
        var xPositions = words.Select(w => w.Bounds.X).OrderBy(x => x).ToList();
        
        // Find leftmost cluster (column 1)
        var leftWords = words.Where(w => w.Bounds.X < xPositions[words.Count / 3]).ToList();
        if (leftWords.Count == 0) return null;

        var col1MinX = leftWords.Min(w => w.Bounds.X);
        var col1MaxX = leftWords.Max(w => w.Bounds.Right);
        var col1MinY = leftWords.Min(w => w.Bounds.Y);
        var col1MaxY = leftWords.Max(w => w.Bounds.Bottom);

        // Find next cluster to the right (column 2)
        var rightWords = words.Where(w => w.Bounds.X > col1MaxX + 20).ToList();
        if (rightWords.Count == 0) return null;

        var col2MinX = rightWords.Min(w => w.Bounds.X);
        var col2MaxX = rightWords.Max(w => w.Bounds.Right);
        var col2MinY = rightWords.Min(w => w.Bounds.Y);
        var col2MaxY = rightWords.Max(w => w.Bounds.Bottom);

        var col1 = new Rectangle(col1MinX, col1MinY, col1MaxX - col1MinX, col1MaxY - col1MinY);
        var col2 = new Rectangle(col2MinX, col2MinY, col2MaxX - col2MinX, col2MaxY - col2MinY);

        return (col1, col2);
    }

    public static IEnumerable<RowCluster> ClusterWordsIntoRows(List<OcrWord> words, int yTol)
    {
        var sorted = words.OrderBy(w => w.Bounds.Y).ThenBy(w => w.Bounds.X).ToList();
        var clusters = new List<RowCluster>();

        foreach (var word in sorted)
        {
            var cluster = clusters.FirstOrDefault(c => Math.Abs(c.Y - word.Bounds.Y) <= yTol);
            if (cluster != null)
            {
                cluster.Words.Add(word);
            }
            else
            {
                clusters.Add(new RowCluster(word.Bounds.Y, new List<OcrWord> { word }));
            }
        }

        return clusters;
    }

    public static IEnumerable<(int PointNumber, string PointName)> ExtractRows(
        List<OcrWord> words,
        (Rectangle col1, Rectangle col2) bands,
        Action<string> log)
    {
        var rows = new List<(int PointNumber, string PointName)>();
        var rowClusters = ClusterWordsIntoRows(words, 10);

        foreach (var cluster in rowClusters)
        {
            var col1Words = cluster.Words.Where(w => Overlaps(w.Bounds, bands.col1)).ToList();
            var col2Words = cluster.Words.Where(w => Overlaps(w.Bounds, bands.col2)).OrderBy(w => w.Bounds.X).ToList();

            if (col1Words.Count == 0 || col2Words.Count == 0) continue;

            // Extract point number from column 1
            var numText = col1Words.FirstOrDefault(w => IsNumeric(w.Text))?.Text;
            if (numText == null) continue;

            if (!int.TryParse(Normalize(numText), out var pointNumber)) continue;

            // Extract point name from column 2
            var pointName = string.Join(" ", col2Words.Select(w => Normalize(w.Text)));
            if (string.IsNullOrWhiteSpace(pointName)) continue;

            // Skip header rows
            if (pointName.ToUpper().Contains("POINT") && pointName.ToUpper().Contains("NAME")) continue;
            if (pointName.ToUpper().Contains("NUMBER")) continue;

            rows.Add((pointNumber, pointName));
        }

        return rows;
    }

    private static bool Overlaps(Rectangle a, Rectangle b)
    {
        return a.X < b.Right && a.Right > b.X && a.Y < b.Bottom && a.Bottom > b.Y;
    }

    private static bool IsNumeric(string text)
    {
        return Regex.IsMatch(text, @"^\d+$");
    }

    public static string Normalize(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        s = Regex.Replace(s, @"\s+", " ");
        return s.Trim();
    }

    public static (List<int> duplicates, List<int> gaps) ValidateSequence(IEnumerable<int> nums)
    {
        var list = nums.ToList();
        var duplicates = list.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        
        var gaps = new List<int>();
        if (list.Count > 0)
        {
            var min = list.Min();
            var max = list.Max();
            var present = new HashSet<int>(list);
            
            for (int i = min; i <= max; i++)
            {
                if (!present.Contains(i))
                {
                    gaps.Add(i);
                }
            }
        }

        return (duplicates, gaps);
    }

    public static void WriteExcel(string outputXlsxPath, IEnumerable<(int PointNumber, string PointName)> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Points");

        // Headers
        sheet.Cell(1, 1).Value = "Point Number";
        sheet.Cell(1, 2).Value = "Point Name";
        sheet.Row(1).Style.Font.Bold = true;

        // Data
        int rowIdx = 2;
        foreach (var (pointNumber, pointName) in rows)
        {
            sheet.Cell(rowIdx, 1).Value = pointNumber;
            sheet.Cell(rowIdx, 2).Value = pointName;
            rowIdx++;
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
        var app = new App();
        var exitCode = await app.RunAsync(DefaultInputFolder, DefaultOutputFolder, TessDataEnvVarName, TessLangEnvVarName);
        Environment.Exit(exitCode);
    }
}
