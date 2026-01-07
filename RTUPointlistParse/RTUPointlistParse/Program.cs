using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using PdfiumViewer;
using Tesseract;
using ClosedXML.Excel;

namespace RTUPointlistParse;

public record OcrWord(string Text, Rectangle Bounds, float Confidence);
public record TableHeader(Rectangle Bounds, int PageIndex, string SourceFile);
public record RowCluster(int PageIndex, string SourceFile, int Y, List<OcrWord> Words);

public class App
{
    private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
    private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";

    private readonly List<string> _logLines = new();

    private void Log(string message)
    {
        Console.WriteLine(message);
        _logLines.Add($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }

    public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVar, string? tessLangEnvVar)
    {
        try
        {
            Log("=== RTU Point List Parser ===");
            Log($"Input folder: {inputFolder}");
            Log($"Output folder: {outputFolder}");

            if (!Directory.Exists(inputFolder))
            {
                Log($"ERROR: Input folder does not exist: {inputFolder}");
                return 1;
            }

            Directory.CreateDirectory(outputFolder);

            var pdfFiles = GetPdfFiles(inputFolder);
            Log($"Found {pdfFiles.Count} PDF file(s) to process");

            if (pdfFiles.Count == 0)
            {
                Log("No PDF files found.");
                return 1;
            }

            var tessDataPath = Environment.GetEnvironmentVariable(tessDataEnvVar ?? "TESSDATA_DIR");
            var tessLang = Environment.GetEnvironmentVariable(tessLangEnvVar ?? "TESSERACT_LANG") ?? "eng";

            if (string.IsNullOrEmpty(tessDataPath))
            {
                Log($"WARNING: {tessDataEnvVar ?? "TESSDATA_DIR"} environment variable not set. Using default tessdata location.");
                tessDataPath = Path.Combine(Directory.GetCurrentDirectory(), "tessdata");
            }

            if (!Directory.Exists(tessDataPath))
            {
                Log($"WARNING: tessdata folder not found at: {tessDataPath}");
            }

            var allRows = new List<(int PointNumber, string PointName)>();

            using (var engine = new TesseractEngine(tessDataPath, tessLang, EngineMode.Default))
            {
                foreach (var pdfPath in pdfFiles)
                {
                    try
                    {
                        Log($"Processing: {Path.GetFileName(pdfPath)}");
                        var bitmaps = RenderPdfToBitmaps(pdfPath, 300, Log);

                        Log($"  Rendered {bitmaps.Count} page(s) from PDF");

                        for (int pageIndex = 0; pageIndex < bitmaps.Count; pageIndex++)
                        {
                            using (var bmp = bitmaps[pageIndex])
                            {
                                var words = OcrWordsFromBitmap(bmp, engine);
                                Log($"  Page {pageIndex + 1}: Extracted {words.Count} words via OCR");

                                var headers = DetectPointNumberHeaders(words);
                                Log($"  Page {pageIndex + 1}: Detected {headers.Count} 'Point Number' header(s)");

                                if (headers.Count > 0)
                                {
                                    var rows = ExtractRowsFromHeaders(words, headers, Log);
                                    allRows.AddRange(rows);
                                    Log($"  Page {pageIndex + 1}: Extracted {rows.Count()} data row(s)");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR processing {Path.GetFileName(pdfPath)}: {ex.Message}");
                    }
                }
            }

            Log($"Total rows extracted: {allRows.Count}");

            var sortedRows = allRows.OrderBy(r => r.PointNumber).ToList();
            Log($"Rows sorted by Point Number");

            var pointNumbers = sortedRows.Select(r => r.PointNumber).ToList();
            var (duplicates, gaps) = ValidateSequence(pointNumbers);

            if (duplicates.Count > 0)
            {
                Log($"WARNING: Found {duplicates.Count} duplicate point number(s): {string.Join(", ", duplicates)}");
            }

            if (gaps.Count > 0)
            {
                Log($"WARNING: Found {gaps.Count} gap(s) in sequence: {string.Join(", ", gaps)}");
            }

            var outputXlsxPath = Path.Combine(outputFolder, "PointList.xlsx");
            WriteExcel(outputXlsxPath, sortedRows);
            Log($"Excel file written: {outputXlsxPath}");

            WriteLog(outputFolder, _logLines);
            Log($"Log file written: {Path.Combine(outputFolder, "PointList.log.txt")}");

            Log("Processing complete.");
            return 0;
        }
        catch (Exception ex)
        {
            Log($"FATAL ERROR: {ex}");
            WriteLog(outputFolder, _logLines);
            return 1;
        }
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
            using (var document = PdfDocument.Load(pdfPath))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    var image = document.Render(i, dpi, dpi, PdfRenderFlags.Annotations);
                    var bmp = new Bitmap(image);
                    bitmaps.Add(bmp);
                    image.Dispose();
                }
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
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                using (var pix = Pix.LoadFromMemory(ms.ToArray()))
                using (var page = engine.Process(pix, PageSegMode.Auto))
                {
                    using (var iter = page.GetIterator())
                    {
                        iter.Begin();
                        do
                        {
                            if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var bounds))
                            {
                                var text = iter.GetText(PageIteratorLevel.Word);
                                var confidence = iter.GetConfidence(PageIteratorLevel.Word);
                                if (!string.IsNullOrWhiteSpace(text))
                                {
                                    var rect = new Rectangle(bounds.X1, bounds.Y1, bounds.Width, bounds.Height);
                                    words.Add(new OcrWord(text.Trim(), rect, confidence));
                                }
                            }
                        } while (iter.Next(PageIteratorLevel.Word));
                    }
                }
            }
        }
        catch (Exception)
        {
            // Silently handle OCR errors
        }
        return words;
    }

    public static List<TableHeader> DetectPointNumberHeaders(List<OcrWord> words)
    {
        var headers = new List<TableHeader>();
        
        for (int i = 0; i < words.Count; i++)
        {
            var word = words[i];
            if (NormalizeText(word.Text).Equals("POINT", StringComparison.OrdinalIgnoreCase))
            {
                for (int j = i + 1; j < words.Count; j++)
                {
                    var nextWord = words[j];
                    if (NormalizeText(nextWord.Text).Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                    {
                        var xOverlap = Math.Min(word.Bounds.Right, nextWord.Bounds.Right) - Math.Max(word.Bounds.Left, nextWord.Bounds.Left);
                        if (xOverlap > 0 && nextWord.Bounds.Top > word.Bounds.Bottom)
                        {
                            var combinedBounds = Rectangle.Union(word.Bounds, nextWord.Bounds);
                            headers.Add(new TableHeader(combinedBounds, 0, ""));
                            break;
                        }
                    }
                }
            }
        }
        
        return headers;
    }

    public static IEnumerable<RowCluster> ClusterWordsIntoRows(List<OcrWord> words, int yTolerance)
    {
        var clusters = new Dictionary<int, List<OcrWord>>();
        
        foreach (var word in words)
        {
            int yKey = word.Bounds.Y / yTolerance * yTolerance;
            
            if (!clusters.ContainsKey(yKey))
            {
                clusters[yKey] = new List<OcrWord>();
            }
            clusters[yKey].Add(word);
        }
        
        return clusters.OrderBy(kvp => kvp.Key)
            .Select(kvp => new RowCluster(0, "", kvp.Key, kvp.Value.OrderBy(w => w.Bounds.X).ToList()));
    }

    public static IEnumerable<(int PointNumber, string PointName)> ExtractRowsFromHeaders(
        List<OcrWord> words, 
        List<TableHeader> headers, 
        Action<string> log)
    {
        var rows = new List<(int PointNumber, string PointName)>();
        
        foreach (var header in headers)
        {
            var belowHeaderWords = words
                .Where(w => w.Bounds.Top > header.Bounds.Bottom && w.Bounds.Bottom < header.Bounds.Bottom + 2000)
                .ToList();
            
            var rowClusters = ClusterWordsIntoRows(belowHeaderWords, 15);
            
            foreach (var cluster in rowClusters)
            {
                var numberWord = cluster.Words
                    .FirstOrDefault(w => w.Bounds.X >= header.Bounds.Left - 50 && 
                                        w.Bounds.X <= header.Bounds.Right + 50 &&
                                        int.TryParse(NormalizeText(w.Text), out _));
                
                if (numberWord != null && int.TryParse(NormalizeText(numberWord.Text), out int pointNumber))
                {
                    var nameWords = cluster.Words
                        .Where(w => w.Bounds.X > numberWord.Bounds.Right)
                        .OrderBy(w => w.Bounds.X)
                        .Take(15)
                        .ToList();
                    
                    if (nameWords.Count > 0)
                    {
                        var pointName = string.Join(" ", nameWords.Select(w => NormalizeText(w.Text)));
                        
                        if (!string.IsNullOrWhiteSpace(pointName) && 
                            !pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase) &&
                            pointName.Length > 2)
                        {
                            rows.Add((pointNumber, pointName));
                        }
                    }
                }
            }
        }
        
        return rows;
    }

    public static string NormalizeText(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        s = Regex.Replace(s, @"\s+", " ");
        s = s.Trim();
        return s;
    }

    public static (List<int> duplicates, List<int> gaps) ValidateSequence(IEnumerable<int> nums)
    {
        var numList = nums.ToList();
        var duplicates = new List<int>();
        var gaps = new List<int>();
        
        if (numList.Count == 0) return (duplicates, gaps);
        
        var seen = new HashSet<int>();
        foreach (var num in numList)
        {
            if (seen.Contains(num))
            {
                if (!duplicates.Contains(num))
                {
                    duplicates.Add(num);
                }
            }
            seen.Add(num);
        }
        
        var uniqueSorted = numList.Distinct().OrderBy(n => n).ToList();
        if (uniqueSorted.Count > 0)
        {
            for (int i = 1; i <= uniqueSorted[uniqueSorted.Count - 1]; i++)
            {
                if (!uniqueSorted.Contains(i))
                {
                    gaps.Add(i);
                }
            }
        }
        
        return (duplicates, gaps);
    }

    public static void WriteExcel(string outputXlsxPath, IEnumerable<(int PointNumber, string PointName)> rows)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Points");
            
            worksheet.Cell(1, 1).Value = "Point Number";
            worksheet.Cell(1, 2).Value = "Point Name";
            worksheet.Row(1).Style.Font.Bold = true;
            
            int rowIndex = 2;
            foreach (var row in rows)
            {
                worksheet.Cell(rowIndex, 1).Value = row.PointNumber;
                worksheet.Cell(rowIndex, 2).Value = row.PointName;
                rowIndex++;
            }
            
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(outputXlsxPath);
        }
    }

    public static void WriteLog(string outputFolder, IEnumerable<string> logLines)
    {
        var logPath = Path.Combine(outputFolder, "PointList.log.txt");
        File.WriteAllLines(logPath, logLines);
    }
}

public class Program
{
    private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
    private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";

    public static async Task<int> Main(string[] args)
    {
        return await new App().RunAsync(DefaultInputFolder, DefaultOutputFolder, "TESSDATA_DIR", "TESSERACT_LANG");
    }
}
