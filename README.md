# RTU Point List Parser

A C# .NET solution for processing RTU (Remote Terminal Unit) point list PDF files and generating structured output.

## Overview

This repository contains tools and applications for parsing and processing point list data from PDF documents. The primary component is the **PointListProcessor** console application, which automates the extraction and formatting of point list information.

## Repository Structure

```
RTUPointListParser/
├── PointListProcessor/         # Main console application
│   ├── Program.cs              # Application logic
│   ├── Point.cs                # Data model
│   ├── PointListProcessor.csproj
│   └── README.md               # Detailed application documentation
│
├── ExamplePointlists/          # Test cases and examples
│   └── Example1/
│       ├── Input/              # Sample input PDFs
│       ├── Expected Output/    # Reference output files
│       └── TestOutput/         # AI-generated test outputs
│
├── Demo/                       # Demonstration files
│   ├── Input/                  # Demo input files
│   └── Output/                 # Demo output files (generated)
│
└── README.md                   # This file
```

## Components

### PointListProcessor Application

A console application that:
- Extracts text from PDF files using the PdfPig library
- Parses point list data into structured format
- Generates formatted output files
- Supports command-line arguments and interactive prompts
- Includes comprehensive error handling and logging

**Key Features:**
- Command-line interface with optional arguments
- Default folder support (Input/TestOutput)
- Automatic output folder creation
- Batch processing of multiple PDF files
- Detailed console logging of processing status

For detailed documentation, see [PointListProcessor/README.md](PointListProcessor/README.md)

## Quick Start

### Prerequisites

- .NET 8.0 SDK or later
- PDF files with extractable text (not scanned images)

### Building the Application

```bash
cd PointListProcessor
dotnet build
```

### Running the Application

#### Interactive Mode

```bash
cd PointListProcessor
dotnet run
```

When prompted, press Enter to use default folders or specify custom paths.

#### Command-Line Mode

```bash
cd PointListProcessor
dotnet run -- --input "/path/to/input" --output "/path/to/output"
```

#### Using the Compiled Executable

```bash
cd PointListProcessor/bin/Debug/net8.0
./PointListProcessor --input "../../Input" --output "../../TestOutput"
```

## Example Usage

### Process Files in Default Folders

```bash
# Place PDF files in PointListProcessor/bin/Debug/net8.0/Input/
cd PointListProcessor
dotnet run
# Press Enter twice to use defaults
# Output files will be in PointListProcessor/bin/Debug/net8.0/TestOutput/
```

### Process Files with Custom Paths

```bash
cd PointListProcessor
dotnet run -- --input "../ExamplePointlists/Example1/Input" --output "../ExamplePointlists/Example1/TestOutput"
```

## Output Format

The application generates text files with the naming pattern: `<InputFileName>_output.txt`

Example output structure:

```
=== Point List Output ===
Generated: 2024-01-06 12:30:45
Total Points: 10

--- Point Details ---

Point #1:
  Name: AI_001
  Type: analog input
  Address: 0
  Description: Temperature Sensor 1

...

--- End of Point List ---
```

## Important Notes

### PDF Format Requirements

The application requires PDF files with **extractable text content**. It does NOT support:
- Scanned PDFs (image-based) without OCR text layer
- PDFs with complex layouts or embedded images only

The sample PDFs in `ExamplePointlists/Example1/Input/` appear to be scanned images and may not extract text successfully. For testing, you'll need PDFs with actual text content.

### Parser Customization

The built-in parser is generic and looks for common patterns in point list data. For specific PDF formats, you may need to customize:
- `ParsePointList()` - Main parsing logic
- `ContainsPointData()` - Line identification logic
- `ParsePointFromLine()` - Field extraction logic
- `GenerateOutput()` - Output format

## Development

### Project Details

- **Language**: C# 12
- **Framework**: .NET 8.0
- **Dependencies**:
  - UglyToad.PdfPig (PDF text extraction)

### Building for Release

```bash
cd PointListProcessor
dotnet publish -c Release -r win-x64 --self-contained
```

Replace `win-x64` with your target platform:
- `win-x64` - Windows 64-bit
- `linux-x64` - Linux 64-bit  
- `osx-x64` - macOS 64-bit
- `osx-arm64` - macOS ARM64 (Apple Silicon)

## Testing

The `ExamplePointlists` folder contains test cases for validation:

1. Place input files in `Example1/Input/`
2. Run the processor to generate output in `Example1/TestOutput/`
3. Compare with reference output in `Example1/Expected Output/`

## Known Limitations

- Only processes PDF files (not other formats like Excel, CSV, etc.)
- Requires PDFs with extractable text (not scanned images)
- Parser may need customization for specific point list formats
- No OCR capability for image-based PDFs

## Future Enhancements

Potential improvements:
- OCR support for scanned PDFs (using Tesseract or similar)
- Support for Excel input files
- Customizable output formats (CSV, JSON, XML)
- Configuration file for parser settings
- GUI application version
- Unit tests for parsing logic

## Contributing

For bug reports, feature requests, or contributions, please open an issue or pull request.

## License

This project is provided as-is for processing RTU point list data.
