# PointListProcessor

A C# .NET Console Application for processing point list PDF files and generating structured output.

## Overview

PointListProcessor is designed to automate the extraction and parsing of point list data from PDF files. The application:
- Extracts text content from PDF files using PdfPig library
- Parses point list data into a structured format
- Generates formatted output files for each processed PDF

## Features

- **Command-line interface**: Accept input and output folder paths via CLI arguments or interactive prompts
- **Default folders**: Use "Input" and "TestOutput" as default folders (relative to executable) when user presses Enter
- **Automatic folder creation**: Creates the output folder if it doesn't exist
- **PDF processing**: Enumerates all PDF files in the input folder and processes them sequentially
- **Error handling**: Validates folder paths, handles missing files gracefully, and logs errors to console
- **Structured output**: Generates "_output.txt" files with the same base name as input files

## Requirements

- .NET 8.0 SDK or later
- UglyToad.PdfPig NuGet package (for PDF text extraction)

## Usage

### Interactive Mode

Run the application without arguments to be prompted for folder paths:

```bash
dotnet run
```

Or run the compiled executable:

```bash
./PointListProcessor
```

When prompted:
- Enter input folder path, or press **Enter** to use default "Input" folder
- Enter output folder path, or press **Enter** to use default "TestOutput" folder

### Command-Line Arguments

Specify input and output folders directly:

```bash
dotnet run -- --input "C:\MyInput" --output "C:\MyOutput"
```

Or with the compiled executable:

```bash
./PointListProcessor --input "/path/to/input" --output "/path/to/output"
```

**Supported arguments:**
- `--input` or `-i`: Specify the input folder path
- `--output` or `-o`: Specify the output folder path

## Building the Application

### Build

```bash
dotnet build
```

### Run

```bash
dotnet run
```

### Publish (for distribution)

```bash
dotnet publish -c Release -r win-x64 --self-contained
```

Replace `win-x64` with your target runtime identifier:
- `win-x64` - Windows 64-bit
- `linux-x64` - Linux 64-bit
- `osx-x64` - macOS 64-bit

## Project Structure

```
PointListProcessor/
├── Program.cs              # Main application logic
├── Point.cs                # Point data model
├── PointListProcessor.csproj  # Project file
└── README.md              # This file
```

## Components

### Program.cs

Contains the main application logic with the following key methods:

- `Main()`: Entry point, handles user input and orchestrates processing
- `GetInputFolder()`: Retrieves input folder from CLI args or user prompt
- `GetOutputFolder()`: Retrieves output folder from CLI args or user prompt
- `ValidateAndPrepareFolders()`: Validates input folder exists and creates output folder
- `ProcessPdfFiles()`: Enumerates and processes all PDF files
- `ExtractTextFromPdf()`: Extracts text content from PDF using PdfPig
- `ParsePointList()`: Parses point list data from extracted text
- `GenerateOutput()`: Generates formatted output string from parsed points

### Point.cs

Data model representing a point with properties:
- `Name`: Point name or identifier
- `Type`: Point type (e.g., AI, AO, DI, DO)
- `Address`: Point address or index
- `Description`: Point description
- `AdditionalProperties`: Dictionary for additional metadata

## Output Format

Generated output files follow this structure:

```
=== Point List Output ===
Generated: 2024-01-06 12:30:45
Total Points: 25

--- Point Details ---

Point #1:
  Name: AI_001
  Type: Analog Input
  Address: 0
  Description: Temperature Sensor

Point #2:
  Name: DI_001
  Type: Digital Input
  Address: 1
  Description: Status Indicator

...

--- End of Point List ---
```

## Error Handling

The application includes comprehensive error handling:

- **Missing input folder**: Displays error and exits
- **Invalid PDF files**: Logs warning and skips to next file
- **No extractable text**: Logs warning and skips file
- **Parsing errors**: Logs warning and continues processing
- **File write errors**: Displays error for specific file and continues

## Limitations

- **PDF Format**: Currently supports PDFs with extractable text. Scanned PDFs (images) without OCR text layer will not extract content.
- **Parser**: The point list parser is generic and may need customization based on specific PDF formats.
- **File Types**: Currently only supports PDF files. Other formats would require additional libraries.

## Customization

To adapt the parser to your specific point list format:

1. Modify `ParsePointList()` method to match your data structure
2. Update `ContainsPointData()` to identify relevant lines
3. Adjust `ParsePointFromLine()` to extract fields correctly
4. Customize `GenerateOutput()` for your desired output format

## License

This project is part of the RTUPointListParser repository.

## Contributing

For bug reports and feature requests, please open an issue in the GitHub repository.
