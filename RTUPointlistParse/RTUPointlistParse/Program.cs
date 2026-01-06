using RTUPointlistParse.Services;

namespace RTUPointlistParse;

class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine("RTU Point List Processor");
        Console.WriteLine("========================");
        Console.WriteLine();

        try
        {
            // Parse command-line arguments
            var (inputPath, outputPath, outputFormat) = ParseArguments(args);

            // Validate and create directories
            if (!Directory.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input directory does not exist: {inputPath}");
                return 1;
            }

            if (!Directory.Exists(outputPath))
            {
                Console.WriteLine($"Creating output directory: {outputPath}");
                Directory.CreateDirectory(outputPath);
            }

            Console.WriteLine($"Input folder:  {Path.GetFullPath(inputPath)}");
            Console.WriteLine($"Output folder: {Path.GetFullPath(outputPath)}");
            Console.WriteLine($"Output format: {outputFormat}");
            Console.WriteLine();

            // Process files
            var result = ProcessFiles(inputPath, outputPath, outputFormat);

            Console.WriteLine();
            Console.WriteLine("Processing Summary:");
            Console.WriteLine($"  Files processed: {result.ProcessedCount}");
            Console.WriteLine($"  Files skipped:   {result.SkippedCount}");
            Console.WriteLine($"  Errors:          {result.ErrorCount}");
            Console.WriteLine();

            if (result.ErrorCount > 0)
            {
                Console.WriteLine("Processing completed with errors.");
                return 1;
            }

            Console.WriteLine("Processing completed successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
            return 1;
        }
    }

    static (string inputPath, string outputPath, string outputFormat) ParseArguments(string[] args)
    {
        string? inputPath = null;
        string? outputPath = null;
        string outputFormat = "txt";

        // Parse command-line arguments
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--input" && i + 1 < args.Length)
            {
                inputPath = args[i + 1];
                i++;
            }
            else if (args[i] == "--output" && i + 1 < args.Length)
            {
                outputPath = args[i + 1];
                i++;
            }
            else if (args[i] == "--format" && i + 1 < args.Length)
            {
                outputFormat = args[i + 1].ToLowerInvariant();
                i++;
            }
            else if (args[i] == "--help" || args[i] == "-h")
            {
                ShowHelp();
                Environment.Exit(0);
            }
        }

        // Interactive prompts if arguments not provided
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            Console.Write("Enter input folder path (Enter for default './Input'): ");
            var userInput = Console.ReadLine();
            inputPath = string.IsNullOrWhiteSpace(userInput) ? "./Input" : userInput;
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            Console.Write("Enter output folder path (Enter for default './TestOutput'): ");
            var userInput = Console.ReadLine();
            outputPath = string.IsNullOrWhiteSpace(userInput) ? "./TestOutput" : userInput;
        }

        return (inputPath, outputPath, outputFormat);
    }

    static void ShowHelp()
    {
        Console.WriteLine("RTU Point List Processor");
        Console.WriteLine("========================");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  RTUPointlistParse [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --input <path>    Input folder path containing point list files");
        Console.WriteLine("                    Default: ./Input");
        Console.WriteLine();
        Console.WriteLine("  --output <path>   Output folder path for generated files");
        Console.WriteLine("                    Default: ./TestOutput");
        Console.WriteLine();
        Console.WriteLine("  --format <type>   Output format: txt, csv, or json");
        Console.WriteLine("                    Default: txt");
        Console.WriteLine();
        Console.WriteLine("  --help, -h        Show this help message");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  RTUPointlistParse --input \"./Input\" --output \"./Output\"");
        Console.WriteLine("  RTUPointlistParse --input \"C:\\Data\\Input\" --output \"C:\\Data\\Output\" --format csv");
        Console.WriteLine();
        Console.WriteLine("Supported input formats:");
        Console.WriteLine("  - PDF files (*.pdf) - text-based only; image-based PDFs require OCR");
        Console.WriteLine("  - Text files (*.txt)");
        Console.WriteLine("  - CSV files (*.csv)");
    }

    static ProcessingResult ProcessFiles(string inputPath, string outputPath, string outputFormat)
    {
        var result = new ProcessingResult();
        var pdfExtractor = new PdfTextExtractor();
        var parser = new PointListParser();
        var writer = new OutputWriter();

        // Discover files
        var pdfFiles = Directory.GetFiles(inputPath, "*.pdf", SearchOption.TopDirectoryOnly)
            .Where(f => !IsHiddenOrSystemFile(f) && new FileInfo(f).Length > 0)
            .ToList();

        var txtFiles = Directory.GetFiles(inputPath, "*.txt", SearchOption.TopDirectoryOnly)
            .Where(f => !IsHiddenOrSystemFile(f) && new FileInfo(f).Length > 0)
            .ToList();

        var csvFiles = Directory.GetFiles(inputPath, "*.csv", SearchOption.TopDirectoryOnly)
            .Where(f => !IsHiddenOrSystemFile(f) && new FileInfo(f).Length > 0)
            .ToList();

        var allFiles = new List<string>();
        allFiles.AddRange(pdfFiles);
        allFiles.AddRange(txtFiles);
        allFiles.AddRange(csvFiles);

        if (allFiles.Count == 0)
        {
            Console.WriteLine($"No files found in: {inputPath}");
            Console.WriteLine("Supported formats: *.pdf, *.txt, *.csv");
            return result;
        }

        Console.WriteLine($"Found {allFiles.Count} file(s) to process:");
        Console.WriteLine($"  PDFs: {pdfFiles.Count}");
        Console.WriteLine($"  Text files: {txtFiles.Count}");
        Console.WriteLine($"  CSV files: {csvFiles.Count}");
        Console.WriteLine();

        // Process each file
        foreach (var filePath in allFiles)
        {
            ProcessFile(filePath, outputPath, outputFormat, pdfExtractor, parser, writer, result);
        }

        return result;
    }

    static void ProcessFile(
        string filePath,
        string outputPath,
        string outputFormat,
        PdfTextExtractor pdfExtractor,
        PointListParser parser,
        OutputWriter writer,
        ProcessingResult result)
    {
        var fileName = Path.GetFileName(filePath);
        var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();

        Console.WriteLine($"Processing: {fileName}");

        try
        {
            string content;

            // Extract text based on file type
            if (fileExtension == ".pdf")
            {
                // Check if PDF has extractable text
                if (!pdfExtractor.HasExtractableText(filePath))
                {
                    Console.WriteLine($"  Skipped: PDF is image-based and requires OCR");
                    result.SkippedCount++;
                    Console.WriteLine();
                    return;
                }

                content = pdfExtractor.ExtractTextFromPdf(filePath);
            }
            else
            {
                // Read text or CSV files directly
                content = File.ReadAllText(filePath);
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                Console.WriteLine($"  Skipped: File is empty");
                result.SkippedCount++;
                Console.WriteLine();
                return;
            }

            // Parse content
            var records = parser.ParsePointList(content);

            if (records.Count == 0)
            {
                Console.WriteLine($"  Skipped: No valid records found");
                result.SkippedCount++;
                Console.WriteLine();
                return;
            }

            // Generate output
            var outputContent = writer.GenerateOutput(records, outputFormat);

            // Determine output file extension
            var outputExtension = outputFormat switch
            {
                "csv" => ".csv",
                "json" => ".json",
                _ => ".txt"
            };

            // Write output file
            var baseName = Path.GetFileNameWithoutExtension(filePath);
            var outputFileName = $"{baseName}_output{outputExtension}";
            var outputFilePath = Path.Combine(outputPath, outputFileName);

            writer.WriteToFile(outputContent, outputFilePath);

            result.ProcessedCount++;
            Console.WriteLine($"  Success: {records.Count} records extracted");
            Console.WriteLine();
        }
        catch (NotSupportedException ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            result.SkippedCount++;
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            result.ErrorCount++;
            Console.WriteLine();
        }
    }

    static bool IsHiddenOrSystemFile(string filePath)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            return (fileInfo.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden ||
                   (fileInfo.Attributes & FileAttributes.System) == FileAttributes.System;
        }
        catch
        {
            return false;
        }
    }
}

class ProcessingResult
{
    public int ProcessedCount { get; set; }
    public int SkippedCount { get; set; }
    public int ErrorCount { get; set; }
}
