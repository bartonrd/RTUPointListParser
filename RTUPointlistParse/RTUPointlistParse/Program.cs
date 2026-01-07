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
    
    // Column detection constants
    private const int ColumnSeparationPixels = 20;  // Minimum horizontal gap between columns
    private const int LeftColumnThresholdDivisor = 3;  // Use first 1/3 of X positions for left column
    
    // OCR artifact filtering
    private static readonly HashSet<string> OcrNoisePatterns = new(StringComparer.OrdinalIgnoreCase)
    {
        "I", "F", "J", "L", "—", "=", "DI", "or", "|", "[", "]", "(", ")", "{", "}", "_"
    };

    private readonly List<string> logLines = new();

    public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVarName, string? tessLangEnvVarName)
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

        var tessDataPath = Environment.GetEnvironmentVariable(tessDataEnvVarName ?? TessDataEnvVarName);
        var tessLang = Environment.GetEnvironmentVariable(tessLangEnvVarName ?? TessLangEnvVarName) ?? "eng";

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

        // Find words that look like point numbers (numeric)
        var numericWords = words.Where(w => w.Text.Any(char.IsDigit)).ToList();
        if (numericWords.Count < 5) return null;  // Need at least some numeric data
        
        // Group numeric words by X position to find the leftmost numeric column
        var numericXPositions = numericWords.Select(w => w.Bounds.X).OrderBy(x => x).ToList();
        var medianNumericX = numericXPositions[numericXPositions.Count / 2];
        
        // Define column 1 band around the numeric words (point numbers)
        var col1Words = numericWords.Where(w => Math.Abs(w.Bounds.X - medianNumericX) < 50).ToList();
        if (col1Words.Count == 0) return null;

        var col1MinX = col1Words.Min(w => w.Bounds.X);
        var col1MaxX = col1Words.Max(w => w.Bounds.Right);
        var col1MinY = words.Min(w => w.Bounds.Y);
        var col1MaxY = words.Max(w => w.Bounds.Bottom);

        // Find text column to the right (point names) - words that are beyond column 1
        var rightWords = words.Where(w => w.Bounds.X > col1MaxX + ColumnSeparationPixels).ToList();
        if (rightWords.Count == 0) return null;
        
        // Find the main cluster of right words (not outliers)
        var rightXPositions = rightWords.Select(w => w.Bounds.X).OrderBy(x => x).ToList();
        var col2StartX = rightXPositions[Math.Min(rightXPositions.Count / 4, rightXPositions.Count - 1)];
        
        var col2Words = rightWords.Where(w => w.Bounds.X >= col2StartX - 20).ToList();
        if (col2Words.Count == 0) return null;

        var col2MinX = col2Words.Min(w => w.Bounds.X);
        var col2MaxX = col2Words.Max(w => w.Bounds.Right);
        var col2MinY = words.Min(w => w.Bounds.Y);
        var col2MaxY = words.Max(w => w.Bounds.Bottom);

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

            // Extract point number from column 1 - look for numeric words and clean them
            var numText = col1Words
                .Select(w => CleanOCRArtifacts(w.Text))
                .FirstOrDefault(w => !string.IsNullOrWhiteSpace(w) && IsNumeric(w));
            
            if (numText == null) continue;
            if (!int.TryParse(numText, out var pointNumber)) continue;

            // Extract point name from column 2 - filter and clean each word
            var cleanedWords = col2Words
                .Select(w => CleanOCRArtifacts(w.Text))
                .Where(w => !IsOcrNoise(w))
                .ToList();
            
            if (cleanedWords.Count == 0) continue;
            
            var pointName = string.Join(" ", cleanedWords);
            pointName = Normalize(pointName);
            
            if (string.IsNullOrWhiteSpace(pointName)) continue;

            // Skip header rows
            var upperName = pointName.ToUpper();
            if (upperName.Contains("POINT") && upperName.Contains("NAME")) continue;
            if (upperName.Contains("NUMBER")) continue;
            if (upperName.Contains("SPARE")) continue;  // Skip spare rows
            
            // Filter out metadata/reference lines
            if (upperName.Contains("LISTING") || upperName.Contains("CONSTRUCTION") || 
                upperName.Contains("ADDED POINT") || upperName.Contains("SYSTEM") ||
                upperName.Contains("REFERENCE") || upperName.Contains("SAP") ||
                upperName.Contains("PLOT BY") || upperName.StartsWith("RESERVED FOR"))
            {
                continue;
            }

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

    private static string CleanOCRArtifacts(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        
        // Remove common OCR artifacts
        string cleaned = text
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
            .Trim();

        // Remove leading OCR artifacts (lowercase l, f, I before uppercase letters)
        while (cleaned.Length > 1 && (cleaned[0] == 'l' || cleaned[0] == 'f' || cleaned[0] == 'I') && 
               char.IsUpper(cleaned[1]))
        {
            cleaned = cleaned.Substring(1);
        }

        if (cleaned.StartsWith("/") && cleaned.Length > 1)
        {
            cleaned = cleaned.Substring(1);
        }

        // Fix common OCR character confusions
        cleaned = cleaned
            .Replace("I'", "1 ")
            .Replace("Il", "11")
            .Replace("Tn", "11")
            .Replace("T15KV", "115KV")
            .Replace("FTRANS", "TRANS")
            .Replace("TS5KV", "115KV")
            .Replace("IN15KV", "115KV")
            .Replace("N15KV", "115KV")
            .Replace("FINO", "NO")
            .Replace("FNO", "NO")
            .Replace("INO", "NO")
            .Replace("fi1S", "115")
            .Replace("fF", "")
            .Replace("cD", "CD")
            .Replace("1155KV", "115KV")
            .Replace("CSF", "CS")
            .Replace("CBF", "CB");

        return cleaned.Trim();
    }

    private static bool IsOcrNoise(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return true;
        if (text.Length <= 1) return true;  // Single characters are likely noise
        if (OcrNoisePatterns.Contains(text)) return true;
        
        // Check if it's all punctuation or special characters
        if (text.All(c => !char.IsLetterOrDigit(c))) return true;
        
        return false;
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
