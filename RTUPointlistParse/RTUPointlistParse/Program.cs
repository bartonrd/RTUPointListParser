using RTUPointlistParse.Services;

namespace RTUPointlistParse;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== RTU Point List Processor ===\n");

        // Parse command-line arguments and get input/output paths
        var (inputPath, outputPath) = ParseArguments(args);

        // Resolve paths with user prompts if needed
        inputPath = ResolveInputPath(inputPath);
        outputPath = ResolveOutputPath(outputPath);

        Console.WriteLine($"Input folder:  {inputPath}");
        Console.WriteLine($"Output folder: {outputPath}");
        Console.WriteLine();

        // Validate input directory exists
        if (!Directory.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input directory does not exist: {inputPath}");
            Environment.Exit(1);
        }

        // Create output directory if it doesn't exist
        try
        {
            Directory.CreateDirectory(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: Failed to create output directory: {ex.Message}");
            Environment.Exit(1);
        }

        // Initialize services
        var pdfExtractor = new PdfTextExtractor();
        var parser = new PointListParser();
        var writer = new OutputWriter();

        // Discover files
        var files = DiscoverFiles(inputPath);

        if (files.Count == 0)
        {
            Console.WriteLine("No files found to process.");
            return;
        }

        Console.WriteLine($"Found {files.Count} file(s) to process.\n");

        // Process each file
        int processedCount = 0;
        int skippedCount = 0;
        int errorCount = 0;

        foreach (var filePath in files)
        {
            var fileName = Path.GetFileName(filePath);
            Console.WriteLine($"Processing: {fileName}");

            try
            {
                // Extract text
                string content;
                try
                {
                    content = pdfExtractor.ExtractTextFromPdf(filePath);
                }
                catch (NotSupportedException ex)
                {
                    Console.WriteLine($"  Skipped: {ex.Message}");
                    skippedCount++;
                    Console.WriteLine();
                    continue;
                }

                // Parse content
                var records = parser.ParsePointList(content);
                Console.WriteLine($"  Parsed {records.Count} point record(s)");

                // Generate output
                var output = writer.GenerateOutput(records);

                // Write output file
                var baseName = Path.GetFileNameWithoutExtension(filePath);
                var outputFileName = $"{baseName}_output.txt";
                var outputFilePath = Path.Combine(outputPath, outputFileName);

                writer.WriteToFile(output, outputFilePath);
                Console.WriteLine($"  Output written to: {outputFileName}");
                processedCount++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error: {ex.Message}");
                errorCount++;
            }

            Console.WriteLine();
        }

        // Print summary
        Console.WriteLine("=== Processing Complete ===");
        Console.WriteLine($"Files processed: {processedCount}");
        Console.WriteLine($"Files skipped:   {skippedCount}");
        Console.WriteLine($"Errors:          {errorCount}");
    }

    /// <summary>
    /// Parses command-line arguments for --input and --output paths.
    /// </summary>
    static (string? inputPath, string? outputPath) ParseArguments(string[] args)
    {
        string? inputPath = null;
        string? outputPath = null;

        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals("--input", StringComparison.OrdinalIgnoreCase))
            {
                inputPath = args[i + 1];
            }
            else if (args[i].Equals("--output", StringComparison.OrdinalIgnoreCase))
            {
                outputPath = args[i + 1];
            }
        }

        return (inputPath, outputPath);
    }

    /// <summary>
    /// Resolves input path with user prompt if not provided.
    /// </summary>
    static string ResolveInputPath(string? inputPath)
    {
        if (!string.IsNullOrWhiteSpace(inputPath))
            return Path.GetFullPath(inputPath);

        Console.Write("Enter input folder path (Enter for default ./Input): ");
        var userInput = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(userInput))
            return Path.GetFullPath("./Input");

        return Path.GetFullPath(userInput);
    }

    /// <summary>
    /// Resolves output path with user prompt if not provided.
    /// </summary>
    static string ResolveOutputPath(string? outputPath)
    {
        if (!string.IsNullOrWhiteSpace(outputPath))
            return Path.GetFullPath(outputPath);

        Console.Write("Enter output folder path (Enter for default ./TestOutput): ");
        var userInput = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(userInput))
            return Path.GetFullPath("./TestOutput");

        return Path.GetFullPath(userInput);
    }

    /// <summary>
    /// Discovers files to process in the input directory.
    /// </summary>
    static List<string> DiscoverFiles(string inputPath)
    {
        var files = new List<string>();

        try
        {
            // Get all PDF files
            var pdfFiles = Directory.GetFiles(inputPath, "*.pdf", SearchOption.TopDirectoryOnly);
            files.AddRange(pdfFiles);

            // Optional: also support TXT and CSV files
            var txtFiles = Directory.GetFiles(inputPath, "*.txt", SearchOption.TopDirectoryOnly);
            var csvFiles = Directory.GetFiles(inputPath, "*.csv", SearchOption.TopDirectoryOnly);
            files.AddRange(txtFiles);
            files.AddRange(csvFiles);

            // Filter out hidden/system files and zero-byte files
            files = files.Where(f =>
            {
                var fileInfo = new FileInfo(f);
                
                // Skip hidden or system files
                if ((fileInfo.Attributes & FileAttributes.Hidden) != 0 ||
                    (fileInfo.Attributes & FileAttributes.System) != 0)
                    return false;

                // Skip zero-byte files
                if (fileInfo.Length == 0)
                    return false;

                // Skip README files
                if (fileInfo.Name.Equals("ReadMe.txt", StringComparison.OrdinalIgnoreCase))
                    return false;

                return true;
            }).ToList();

            // Sort for consistent processing order
            files.Sort();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Error discovering files: {ex.Message}");
        }

        return files;
    }
}
