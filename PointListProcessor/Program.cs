using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PointListProcessor;

/// <summary>
/// Main program class for PointListProcessor console application.
/// This application processes point list PDF files and generates structured output.
/// </summary>
class Program
{
    /// <summary>
    /// Entry point for the application.
    /// </summary>
    /// <param name="args">Command-line arguments: --input "path" --output "path"</param>
    static void Main(string[] args)
    {
        Console.WriteLine("=== PointListProcessor ===");
        Console.WriteLine("Process point list PDF files and generate structured output.");
        Console.WriteLine();

        // Parse command-line arguments or prompt for input
        string inputFolder = GetInputFolder(args);
        string outputFolder = GetOutputFolder(args);

        Console.WriteLine($"Input folder: {inputFolder}");
        Console.WriteLine($"Output folder: {outputFolder}");
        Console.WriteLine();

        // Validate and create folders
        if (!ValidateAndPrepareFolders(inputFolder, outputFolder))
        {
            Console.WriteLine("Failed to validate or create folders. Exiting.");
            return;
        }

        // Process all PDF files in the input folder
        ProcessPdfFiles(inputFolder, outputFolder);

        Console.WriteLine();
        Console.WriteLine("Processing complete. Press any key to exit...");
        
        // Only wait for keypress if running in interactive mode
        try
        {
            if (!Console.IsInputRedirected)
            {
                Console.ReadKey();
            }
        }
        catch
        {
            // Ignore errors if console is not available
        }
    }

    /// <summary>
    /// Gets the input folder path from command-line arguments or user prompt.
    /// </summary>
    /// <param name="args">Command-line arguments</param>
    /// <returns>Input folder path</returns>
    static string GetInputFolder(string[] args)
    {
        // Check for --input argument
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i] == "--input" || args[i] == "-i")
            {
                return args[i + 1];
            }
        }

        // Prompt user for input
        Console.Write("Enter input folder path (press Enter for default 'Input'): ");
        string? input = Console.ReadLine();
        
        // Use default if empty
        if (string.IsNullOrWhiteSpace(input))
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Input");
        }

        return input;
    }

    /// <summary>
    /// Gets the output folder path from command-line arguments or user prompt.
    /// </summary>
    /// <param name="args">Command-line arguments</param>
    /// <returns>Output folder path</returns>
    static string GetOutputFolder(string[] args)
    {
        // Check for --output argument
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i] == "--output" || args[i] == "-o")
            {
                return args[i + 1];
            }
        }

        // Prompt user for output
        Console.Write("Enter output folder path (press Enter for default 'TestOutput'): ");
        string? output = Console.ReadLine();
        
        // Use default if empty
        if (string.IsNullOrWhiteSpace(output))
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestOutput");
        }

        return output;
    }

    /// <summary>
    /// Validates that the input folder exists and creates the output folder if needed.
    /// </summary>
    /// <param name="inputFolder">Input folder path</param>
    /// <param name="outputFolder">Output folder path</param>
    /// <returns>True if validation succeeds, false otherwise</returns>
    static bool ValidateAndPrepareFolders(string inputFolder, string outputFolder)
    {
        try
        {
            // Validate input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"ERROR: Input folder does not exist: {inputFolder}");
                return false;
            }

            // Create output folder if it doesn't exist
            if (!Directory.Exists(outputFolder))
            {
                Console.WriteLine($"Creating output folder: {outputFolder}");
                Directory.CreateDirectory(outputFolder);
            }

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR validating folders: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Processes all PDF files in the input folder.
    /// </summary>
    /// <param name="inputFolder">Input folder path</param>
    /// <param name="outputFolder">Output folder path</param>
    static void ProcessPdfFiles(string inputFolder, string outputFolder)
    {
        try
        {
            // Get all PDF files from input folder
            string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf", SearchOption.TopDirectoryOnly);

            if (pdfFiles.Length == 0)
            {
                Console.WriteLine("No PDF files found in the input folder.");
                return;
            }

            Console.WriteLine($"Found {pdfFiles.Length} PDF file(s) to process.");
            Console.WriteLine();

            int processed = 0;
            int skipped = 0;

            // Process each PDF file
            foreach (string pdfFile in pdfFiles)
            {
                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(pdfFile)}");

                    // Extract text from PDF
                    string extractedText = ExtractTextFromPdf(pdfFile);

                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        Console.WriteLine($"  WARNING: No text extracted from {Path.GetFileName(pdfFile)}. Skipping.");
                        skipped++;
                        continue;
                    }

                    // Parse point list
                    List<Point> points = ParsePointList(extractedText);

                    if (points.Count == 0)
                    {
                        Console.WriteLine($"  WARNING: No points found in {Path.GetFileName(pdfFile)}. Skipping.");
                        skipped++;
                        continue;
                    }

                    // Generate output
                    string output = GenerateOutput(points);

                    // Save output file
                    string baseFileName = Path.GetFileNameWithoutExtension(pdfFile);
                    string outputFileName = $"{baseFileName}_output.txt";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    File.WriteAllText(outputPath, output);

                    Console.WriteLine($"  SUCCESS: Generated {outputFileName} with {points.Count} point(s).");
                    processed++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ERROR processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                    skipped++;
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Processing summary: {processed} processed, {skipped} skipped.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR enumerating PDF files: {ex.Message}");
        }
    }

    /// <summary>
    /// Extracts text from a PDF file using PdfPig.
    /// </summary>
    /// <param name="filePath">Path to the PDF file</param>
    /// <returns>Extracted text content</returns>
    static string ExtractTextFromPdf(string filePath)
    {
        try
        {
            using (PdfDocument document = PdfDocument.Open(filePath))
            {
                StringBuilder text = new StringBuilder();

                // Extract text from each page
                foreach (Page page in document.GetPages())
                {
                    text.AppendLine($"--- Page {page.Number} ---");
                    text.AppendLine(page.Text);
                    text.AppendLine();
                }

                return text.ToString();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to extract text from PDF: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Parses point list data from extracted text content.
    /// This is a basic parser that looks for common point list patterns.
    /// </summary>
    /// <param name="content">Extracted text content</param>
    /// <returns>List of parsed Point objects</returns>
    static List<Point> ParsePointList(string content)
    {
        List<Point> points = new List<Point>();

        try
        {
            // Split content into lines
            string[] lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Basic parsing: look for patterns that might indicate point data
            // This is a simplified parser - real implementation would need to be customized
            // based on the actual format of the point list PDFs

            foreach (string line in lines)
            {
                // Skip header lines or empty lines
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("---"))
                    continue;

                // Look for lines that contain typical point information
                // Pattern: Look for lines with multiple fields separated by spaces or tabs
                string trimmedLine = line.Trim();
                
                // Simple heuristic: If line has multiple words and contains numbers or common keywords
                if (ContainsPointData(trimmedLine))
                {
                    Point? point = ParsePointFromLine(trimmedLine);
                    if (point != null)
                    {
                        points.Add(point);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  WARNING: Error during parsing: {ex.Message}");
        }

        return points;
    }

    /// <summary>
    /// Checks if a line contains potential point data.
    /// </summary>
    /// <param name="line">Line to check</param>
    /// <returns>True if line appears to contain point data</returns>
    static bool ContainsPointData(string line)
    {
        // Look for common point list indicators
        string[] pointKeywords = { "AI", "AO", "DI", "DO", "BI", "BO", "analog", "digital", "input", "output", "point" };
        
        string lowerLine = line.ToLower();
        
        // Check if line contains point-related keywords
        foreach (string keyword in pointKeywords)
        {
            if (lowerLine.Contains(keyword))
                return true;
        }

        // Check if line has a structured format (multiple fields)
        string[] fields = line.Split(new[] { ' ', '\t', '|', ',' }, StringSplitOptions.RemoveEmptyEntries);
        
        // If line has 3+ fields and contains at least one number, it might be point data
        if (fields.Length >= 3 && fields.Any(f => Regex.IsMatch(f, @"\d+")))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Parses a Point object from a line of text.
    /// </summary>
    /// <param name="line">Line to parse</param>
    /// <returns>Parsed Point object or null if parsing fails</returns>
    static Point? ParsePointFromLine(string line)
    {
        try
        {
            // Split line into fields
            string[] fields = line.Split(new[] { ' ', '\t', '|', ',' }, StringSplitOptions.RemoveEmptyEntries);

            if (fields.Length < 2)
                return null;

            Point point = new Point
            {
                Name = fields[0],
                Type = fields.Length > 1 ? fields[1] : "",
                Address = fields.Length > 2 ? fields[2] : "",
                Description = fields.Length > 3 ? string.Join(" ", fields.Skip(3)) : ""
            };

            return point;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Generates output string from a list of Point objects.
    /// </summary>
    /// <param name="points">List of Point objects</param>
    /// <returns>Formatted output string</returns>
    static string GenerateOutput(List<Point> points)
    {
        StringBuilder output = new StringBuilder();

        // Add header
        output.AppendLine("=== Point List Output ===");
        output.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        output.AppendLine($"Total Points: {points.Count}");
        output.AppendLine();
        output.AppendLine("--- Point Details ---");
        output.AppendLine();

        // Format each point
        int index = 1;
        foreach (Point point in points)
        {
            output.AppendLine($"Point #{index}:");
            output.AppendLine($"  Name: {point.Name}");
            output.AppendLine($"  Type: {point.Type}");
            output.AppendLine($"  Address: {point.Address}");
            output.AppendLine($"  Description: {point.Description}");
            
            if (point.AdditionalProperties.Any())
            {
                output.AppendLine("  Additional Properties:");
                foreach (var prop in point.AdditionalProperties)
                {
                    output.AppendLine($"    {prop.Key}: {prop.Value}");
                }
            }
            
            output.AppendLine();
            index++;
        }

        output.AppendLine("--- End of Point List ---");

        return output.ToString();
    }
}
