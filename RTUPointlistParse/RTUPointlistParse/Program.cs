using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using PdfiumViewer;
using Tesseract;
using ClosedXML.Excel;

namespace RTUPointlistParse;

record OcrWord(string Text, Rectangle Bounds, float Confidence);
record RowCluster(int Y, List<OcrWord> Words);

class App
{
    private const string DefaultInputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\Input";
    private const string DefaultOutputFolder = "C:\\dev\\RTUPointListParser\\ExamplePointlists\\Example1\\TestOutput";
    private readonly List<string> logLines = new();

    public static async Task<int> Main(string[] args)
    {
        var app = new App();
        return await app.RunAsync(DefaultInputFolder, DefaultOutputFolder, "TESSDATA_DIR", "TESSERACT_LANG");
    }

    public async Task<int> RunAsync(string inputFolder, string outputFolder, string? tessDataEnvVar, string? tessLangEnvVar)
    {
        await Task.CompletedTask; // Make it async as required
        
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
        Log($"Found {pdfFiles.Count} PDF file(s) to process");
        Log("");

        if (pdfFiles.Count == 0)
        {
            Log("No PDF files found in the input folder.");
            WriteLog(outputFolder, logLines);
            return 0;
        }

        var tessDataPath = Environment.GetEnvironmentVariable(tessDataEnvVar ?? "TESSDATA_DIR");
        var tessLang = Environment.GetEnvironmentVariable(tessLangEnvVar ?? "TESSERACT_LANG") ?? "eng";

        if (string.IsNullOrEmpty(tessDataPath))
        {
            Log($"Warning: {tessDataEnvVar ?? "TESSDATA_DIR"} environment variable not set. Using default tessdata path.");
            tessDataPath = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", "tessdata");
        }

        Log($"Tesseract data path: {tessDataPath}");
        Log($"Tesseract language: {tessLang}");
        Log("");

        var allRows = new List<(int PointNumber, string PointName)>();

        using (var engine = new TesseractEngine(tessDataPath, tessLang, EngineMode.Default))
        {
            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    Log($"Processing: {Path.GetFileName(pdfFile)}");
                    
                    var bitmaps = RenderPdfToBitmaps(pdfFile, 300, Log);
                    Log($"  Rendered {bitmaps.Count} page(s)");

                    int pageNum = 0;
                    foreach (var bitmap in bitmaps)
                    {
                        pageNum++;
                        try
                        {
                            var words = OcrWordsFromBitmap(bitmap, engine);
                            Log($"  Page {pageNum}: OCR found {words.Count} words");

                            var bands = DetectFirstTwoColumnBands(words);
                            if (bands.HasValue)
                            {
                                var rows = ExtractRows(words, bands.Value, Log).ToList();
                                allRows.AddRange(rows);
                                Log($"  Page {pageNum}: Extracted {rows.Count} rows");
                            }
                            else
                            {
                                Log($"  Page {pageNum}: Could not detect table columns");
                            }
                        }
                        finally
                        {
                            bitmap.Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"  Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                }
            }
        }

        Log("");
        Log($"Total rows extracted: {allRows.Count}");

        if (allRows.Count > 0)
        {
            // Sort by Point Number
            allRows = allRows.OrderBy(r => r.PointNumber).ToList();

            // Validate sequence
            var pointNumbers = allRows.Select(r => r.PointNumber).ToList();
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
            var outputXlsxPath = Path.Combine(outputFolder, "PointList.xlsx");
            WriteExcel(outputXlsxPath, allRows);
            Log($"Written output to: {outputXlsxPath}");
        }
        else
        {
            Log("Warning: No rows extracted from PDFs");
        }

        // Write log
        WriteLog(outputFolder, logLines);
        Log("");
        Log("Processing complete.");

        return 0;
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
            using (var document = PdfDocument.Load(pdfPath))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    var image = document.Render(i, dpi, dpi, PdfRenderFlags.Annotations);
                    bitmaps.Add((Bitmap)image);
                }
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
        try
        {
            using (var pix = PixConverter.ToPix(bmp))
            using (var page = engine.Process(pix))
            {
                using (var iter = page.GetIterator())
                {
                    iter.Begin();
                    do
                    {
                        if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var bounds))
                        {
                            var text = iter.GetText(PageIteratorLevel.Word);
                            var conf = iter.GetConfidence(PageIteratorLevel.Word);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                words.Add(new OcrWord(text.Trim(), 
                                    new Rectangle(bounds.X1, bounds.Y1, bounds.Width, bounds.Height), 
                                    conf));
                            }
                        }
                    } while (iter.Next(PageIteratorLevel.Word));
                }
            }
        }
        catch (Exception)
        {
            // Ignore OCR errors
        }
        return words;
    }

    public static (Rectangle col1, Rectangle col2)? DetectFirstTwoColumnBands(List<OcrWord> words)
    {
        if (words.Count < 10) return null;

        // Look for "POINT" and "NUMBER" header words (may be stacked)
        var headerWords = words.Where(w => 
            w.Text.Contains("POINT", StringComparison.OrdinalIgnoreCase) ||
            w.Text.Contains("NUMBER", StringComparison.OrdinalIgnoreCase) ||
            w.Text.Contains("NAME", StringComparison.OrdinalIgnoreCase)).ToList();

        if (headerWords.Count == 0)
        {
            // Fall back to detecting numeric column
            var numericWords = words.Where(w => int.TryParse(w.Text, out _)).ToList();
            if (numericWords.Count < 5) return null;

            // Find leftmost numeric column
            var leftmostX = numericWords.Min(w => w.Bounds.X);
            var col1Words = numericWords.Where(w => Math.Abs(w.Bounds.X - leftmostX) < 50).ToList();
            if (col1Words.Count < 5) return null;

            var col1Left = col1Words.Min(w => w.Bounds.X);
            var col1Right = col1Words.Max(w => w.Bounds.Right);

            // Find next column to the right
            var col2Words = words.Where(w => w.Bounds.X > col1Right + 10).ToList();
            if (col2Words.Count < 5) return null;

            var col2Left = col2Words.Min(w => w.Bounds.X);
            var col2Right = col2Words.Max(w => w.Bounds.Right);

            var topY = Math.Min(col1Words.Min(w => w.Bounds.Y), col2Words.Min(w => w.Bounds.Y));
            var bottomY = Math.Max(col1Words.Max(w => w.Bounds.Bottom), col2Words.Max(w => w.Bounds.Bottom));

            return (new Rectangle(col1Left, topY, col1Right - col1Left, bottomY - topY),
                    new Rectangle(col2Left, topY, col2Right - col2Left, bottomY - topY));
        }

        // Find columns based on header
        var pointNumberHeaders = headerWords.Where(w => 
            w.Text.Contains("POINT", StringComparison.OrdinalIgnoreCase) ||
            w.Text.Contains("NUMBER", StringComparison.OrdinalIgnoreCase)).ToList();

        var pointNameHeaders = headerWords.Where(w => 
            w.Text.Contains("NAME", StringComparison.OrdinalIgnoreCase)).ToList();

        if (pointNumberHeaders.Count == 0) return null;

        // Get column 1 bounds (POINT NUMBER)
        var col1HeaderX = pointNumberHeaders.Min(w => w.Bounds.X);
        var col1Left2 = col1HeaderX;
        
        // Find column 2 bounds (POINT NAME) - next column to the right
        int col2Left2;
        if (pointNameHeaders.Count > 0)
        {
            col2Left2 = pointNameHeaders.Min(w => w.Bounds.X);
        }
        else
        {
            // Estimate based on words to the right
            col2Left2 = col1Left2 + 150;
        }

        // Define column widths
        var col1Width = col2Left2 - col1Left2 - 10;
        var col2Width = words.Max(w => w.Bounds.Right) - col2Left2;

        var topY2 = words.Min(w => w.Bounds.Y);
        var bottomY2 = words.Max(w => w.Bounds.Bottom);

        return (new Rectangle(col1Left2, topY2, col1Width, bottomY2 - topY2),
                new Rectangle(col2Left2, topY2, col2Width, bottomY2 - topY2));
    }

    public static IEnumerable<RowCluster> ClusterWordsIntoRows(List<OcrWord> words, int yTol)
    {
        if (words.Count == 0) yield break;

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

        foreach (var cluster in clusters.OrderBy(c => c.Y))
        {
            yield return cluster;
        }
    }

    public static IEnumerable<(int PointNumber, string PointName)> ExtractRows(
        List<OcrWord> words, 
        (Rectangle col1, Rectangle col2) bands, 
        Action<string> log)
    {
        // Filter words into columns
        var col1Words = words.Where(w => 
            w.Bounds.X >= bands.col1.X && 
            w.Bounds.X < bands.col1.Right).ToList();

        var col2Words = words.Where(w => 
            w.Bounds.X >= bands.col2.X && 
            w.Bounds.X < bands.col2.Right).ToList();

        // Cluster into rows
        var col1Rows = ClusterWordsIntoRows(col1Words, 15).ToList();
        var col2Rows = ClusterWordsIntoRows(col2Words, 15).ToList();

        // Match rows by Y position
        foreach (var row1 in col1Rows)
        {
            // Find first numeric word in column 1
            var numericWord = row1.Words.FirstOrDefault(w => int.TryParse(Normalize(w.Text), out _));
            if (numericWord == null) continue;

            if (!int.TryParse(Normalize(numericWord.Text), out int pointNumber)) continue;

            // Skip header rows (like "POINT NUMBER")
            if (pointNumber == 0 || row1.Words.Any(w => 
                w.Text.Contains("POINT", StringComparison.OrdinalIgnoreCase) ||
                w.Text.Contains("NUMBER", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            // Find matching row in column 2 (within Y tolerance)
            var row2 = col2Rows.FirstOrDefault(r => Math.Abs(r.Y - row1.Y) <= 20);
            if (row2 == null || row2.Words.Count == 0) continue;

            // Combine words in column 2, ordered by X
            var pointNameWords = row2.Words.OrderBy(w => w.Bounds.X).Select(w => Normalize(w.Text));
            var pointName = string.Join(" ", pointNameWords).Trim();

            // Skip empty or invalid names
            if (string.IsNullOrWhiteSpace(pointName)) continue;
            if (pointName.Length < 3) continue;
            if (pointName.Contains("SPARE", StringComparison.OrdinalIgnoreCase)) continue;

            // Skip header-like content
            if (pointName.Contains("POINT", StringComparison.OrdinalIgnoreCase) ||
                pointName.Contains("NAME", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            yield return (pointNumber, pointName);
        }
    }

    public static string Normalize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        
        // Remove common OCR artifacts
        s = s.Trim();
        s = Regex.Replace(s, @"[|_\[\]{}]", " ");
        s = Regex.Replace(s, @"\s+", " ");
        return s.Trim();
    }

    public static (List<int> duplicates, List<int> gaps) ValidateSequence(IEnumerable<int> nums)
    {
        var duplicates = new List<int>();
        var gaps = new List<int>();

        var numsList = nums.OrderBy(n => n).ToList();
        if (numsList.Count == 0) return (duplicates, gaps);

        // Find duplicates
        for (int i = 1; i < numsList.Count; i++)
        {
            if (numsList[i] == numsList[i - 1])
            {
                if (!duplicates.Contains(numsList[i]))
                {
                    duplicates.Add(numsList[i]);
                }
            }
        }

        // Find gaps
        for (int i = 1; i < numsList.Count; i++)
        {
            if (numsList[i] - numsList[i - 1] > 1)
            {
                for (int g = numsList[i - 1] + 1; g < numsList[i]; g++)
                {
                    gaps.Add(g);
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

            // Headers
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

            workbook.SaveAs(outputXlsxPath);
        }
    }

    public static void WriteLog(string outputFolder, IEnumerable<string> logLines)
    {
        var logPath = Path.Combine(outputFolder, "PointList.log.txt");
        File.WriteAllLines(logPath, logLines);
    }
}
