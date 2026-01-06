# RTU Point List Parser

A .NET console application that reads point-list files from an input folder, parses the data, and writes formatted CSV output files to an output folder.

## Features

- **Command-line Arguments**: Specify input/output paths via `--input` and `--output` flags
- **Interactive Prompts**: If arguments are not provided, the app prompts for paths with defaults
- **Multiple File Formats**: Supports PDF (via PdfPig), TXT, and CSV input files
- **Flexible Parsing**: Handles space-separated, tab-separated, and CSV formats
- **Sorted Output**: Generates deterministic CSV output sorted by Name and Address
- **Error Handling**: Graceful handling of missing files, parse errors, and I/O exceptions
- **Comprehensive Logging**: Console logs for processing status, warnings, and errors

## Requirements

- .NET 10.0 or later
- PdfPig NuGet package (for PDF text extraction)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/bartonrd/RTUPointListParser.git
   cd RTUPointListParser
   ```

2. Build the project:
   ```bash
   cd RTUPointlistParse/RTUPointlistParse
   dotnet build
   ```

## Usage

### Command-Line Mode

Specify both input and output paths:
```bash
dotnet run -- --input "C:\Path\To\Input" --output "C:\Path\To\Output"
```

Or use relative paths:
```bash
dotnet run -- --input "./Input" --output "./TestOutput"
```

### Interactive Mode

Run without arguments to be prompted for paths:
```bash
dotnet run
```

Press Enter at each prompt to use the default paths:
- Default input: `./Input`
- Default output: `./TestOutput`

## Input File Formats

### Space/Tab-Separated Format

```
Name            Address     Type        Unit    Description
AI_Voltage_1    40001       AI          V       Primary voltage measurement
AI_Current_1    40002       AI          A       Primary current measurement
DI_Status_1     10001       DI          -       Main breaker status
```

### CSV Format

```
Name,Address,Type,Unit,Description
AI_Voltage_1,40001,AI,V,Primary voltage measurement
AI_Current_1,40002,AI,A,Primary current measurement
DI_Status_1,10001,DI,-,Main breaker status
```

### PDF Format

PDF files are supported via the PdfPig library. Text is extracted from all pages and then parsed.

**Note**: PDF files that are scanned images without embedded text cannot be parsed. These files will produce empty output.

## Output Format

All output files are generated in CSV format with UTF-8 encoding (without BOM):

```csv
Name,Address,Type,Unit,Description
AI_Current_1,40002,AI,A,Primary current measurement
AI_Voltage_1,40001,AI,V,Primary voltage measurement
DI_Status_1,10001,DI,-,Main breaker status
```

Output files are named: `{InputBaseName}_output.txt`

## Point Record Model

Each parsed point record contains:
- **Name**: Point identifier
- **Address**: Register/memory address
- **Type**: Point type (AI, AO, DI, DO, Counter, etc.)
- **Unit**: Unit of measurement (V, A, kW, °F, etc.)
- **Description**: Human-readable description

## Architecture

The application is organized into the following structure:

```
RTUPointlistParse/
├── Program.cs              # Main entry point, CLI handling, orchestration
├── Models/
│   └── PointRecord.cs      # Point record data model
└── Services/
    ├── PdfTextExtractor.cs # PDF text extraction service
    ├── PointListParser.cs  # Point list parsing logic
    └── OutputWriter.cs     # CSV output generation
```

## Examples

Sample input files are provided in the `Input/` directory:
- `sample_pointlist.txt` - Space-separated format example
- `additional_points.csv` - CSV format example

Run with sample files:
```bash
cd RTUPointListParser
dotnet run --project RTUPointlistParse/RTUPointlistParse -- --input "./Input" --output "./TestOutput"
```

## Error Handling

The application handles:
- Missing input directory (exits with error)
- No files found (logs and exits gracefully)
- Unsupported file types (skipped with warning)
- PDF files without extractable text (skipped with message)
- Parse errors (logged with line information)
- File I/O errors (logged and counted in summary)

## Parsing Behavior

- **Header/Footer Detection**: Automatically skips common headers, footers, and title lines
- **Whitespace Normalization**: Robust to inconsistent spacing and tabs
- **Empty Field Handling**: Missing fields are set to empty strings
- **Line Skipping**: Invalid lines are logged but don't stop processing

## Output

At the end of processing, a summary is displayed:

```
=== Processing Complete ===
Files processed: 2
Files skipped:   0
Errors:          0
```

## Future Enhancements

- **Validation Mode**: Add `--compare` flag to diff generated vs. expected output
- **Additional Formats**: Support for Excel, XML, or JSON input
- **Custom Parsers**: Configurable parsing rules per file type
- **OCR Support**: Extract text from scanned PDF images

## License

This project is licensed under the MIT License.

## Contributing

Contributions are welcome! Please submit issues or pull requests on GitHub.
