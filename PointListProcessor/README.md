# PointListProcessor

A C# console application that reads point list files from PDF documents, parses them, and generates Excel output files.

## Features

- Accepts user-specified input and output directories via command-line arguments or interactive prompts
- Defaults to `Input` folder for source files and `TestOutput` folder for generated files
- Reads and parses point list files from PDF documents
- Generates Excel (.xlsx) output files with structured point data
- Error handling and validation for paths and file processing

## Prerequisites

- .NET 8.0 SDK or later

## Building the Application

```bash
cd PointListProcessor
dotnet restore
dotnet build
```

## Usage

### 1. Command-Line with Named Arguments

```bash
dotnet run -- --input "path/to/input" --output "path/to/output"
```

Example:
```bash
dotnet run -- --input "./ExamplePointlists/Example1/Input" --output "./ExamplePointlists/Example1/TestOutput"
```

### 2. Command-Line with Positional Arguments

```bash
dotnet run -- "path/to/input" "path/to/output"
```

Example:
```bash
dotnet run -- "./ExamplePointlists/Example1/Input" "./ExamplePointlists/Example1/TestOutput"
```

### 3. Interactive Mode

Simply run the application without arguments and it will prompt you for input/output paths:

```bash
dotnet run
```

When prompted:
- Press Enter to use the default `Input` folder (relative to the application directory)
- Press Enter again to use the default `TestOutput` folder (relative to the application directory)
- Or provide custom paths

## Output Format

The application generates Excel files with the following structure:
- Point Name
- Point Type
- Description

Output files are named using the pattern: `{original_filename}_output.xlsx`

## Error Handling

- Validates that input folder exists before processing
- Creates output folder if it doesn't exist
- Logs errors for files that cannot be processed
- Continues processing remaining files if one file fails

## Dependencies

- **itext7** (8.0.5) - PDF text extraction
- **EPPlus** (7.5.1) - Excel file generation (Non-Commercial License)

## License

EPPlus is used under the Non-Commercial License. For commercial use, please ensure you have the appropriate license.
