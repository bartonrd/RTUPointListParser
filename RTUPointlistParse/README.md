# RTU Point List Processor

A C# .NET console application that reads point-list files from an input folder, parses the data, and writes formatted output files to an output folder.

## Features

- **Multiple Input Formats**: Supports PDF (text-based), TXT, and CSV files
- **Multiple Output Formats**: Generate output in TXT, CSV, or JSON format
- **Command-Line Interface**: Supports both command-line arguments and interactive prompts
- **Intelligent PDF Handling**: Detects image-based PDFs and provides clear guidance
- **Robust Error Handling**: Gracefully handles missing files, empty content, and parsing errors
- **Flexible Configuration**: Customizable input/output paths and format selection

## Usage

### Command-Line Arguments

```bash
RTUPointlistParse [options]

Options:
  --input <path>    Input folder path containing point list files
                    Default: ./Input

  --output <path>   Output folder path for generated files
                    Default: ./TestOutput

  --format <type>   Output format: txt, csv, or json
                    Default: txt

  --help, -h        Show this help message
```

### Examples

```bash
# Use default paths (./Input and ./TestOutput)
dotnet run

# Specify custom paths
dotnet run -- --input "./Data/Input" --output "./Data/Output"

# Specify output format
dotnet run -- --input "./Input" --output "./Output" --format csv

# Show help
dotnet run -- --help
```

### Interactive Mode

If you don't provide command-line arguments, the application will prompt you for input:

```
Enter input folder path (Enter for default './Input'): 
Enter output folder path (Enter for default './TestOutput'):
```

Press Enter at each prompt to use the default values.

## Supported Input Formats

- **PDF files (*.pdf)**: Text-based PDFs only. Image-based PDFs require OCR (Optical Character Recognition).
- **Text files (*.txt)**: Plain text point list files.
- **CSV files (*.csv)**: Comma-separated value files.

## Output Formats

### TXT Format
A human-readable text format with each point record displayed with labeled fields.

### CSV Format
A comma-separated values format suitable for importing into spreadsheet applications.

### JSON Format
A structured JSON format for programmatic processing.

## Building the Project

```bash
cd RTUPointlistParse/RTUPointlistParse
dotnet build
```

## Running the Application

```bash
cd RTUPointlistParse/RTUPointlistParse
dotnet run -- --input "../../Input" --output "../../TestOutput"
```

## Project Structure

```
RTUPointlistParse/
├── Models/
│   └── PointRecord.cs          # Data model for point records
├── Services/
│   ├── PdfTextExtractor.cs     # PDF text extraction service
│   ├── PointListParser.cs      # Point list parsing service
│   └── OutputWriter.cs         # Output generation service
└── Program.cs                  # Main application entry point
```

## Dependencies

- **UglyToad.PdfPig**: For PDF text extraction
- **.NET 10.0**: Target framework

## Known Limitations

1. **Image-based PDFs**: The application cannot extract text from scanned/image-based PDFs. These require OCR software. The application will detect this and skip such files with a clear message.

2. **Point List Format**: The parser includes basic logic for common delimiters (tabs, pipes, commas). For specific custom formats, the `ParseLine` method in `PointListParser.cs` may need customization.

3. **Header Detection**: The application attempts to skip common headers/footers but may need adjustment for specific formats.

## Error Handling

The application provides clear error messages for common scenarios:

- Missing input directory
- No files found in input directory
- Empty files
- Image-based PDFs requiring OCR
- Parse failures (with line numbers when available)

## Output File Naming

For each input file, the application generates an output file with the pattern:
- Input: `filename.ext`
- Output: `filename_output.{txt|csv|json}`

## License

See repository license for details.
