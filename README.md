# RTU Point List Parser

A C# .NET 8 Windows Console Application for extracting point list data from image-based PDF files using OCR and generating Excel output.

## Overview

This tool processes scanned PDF documents containing RTU point list tables. It uses OCR (Optical Character Recognition) to extract text, detects table structures by analyzing word geometry, and exports the first two columns (Point Number and Point Name) to an Excel file.

## Requirements

- **Operating System**: Windows (required for PdfiumViewer native libraries)
- **.NET**: .NET 8.0 SDK or later
- **Tesseract Data**: `tessdata` directory with English language data (`eng.traineddata`)

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/bartonrd/RTUPointListParser.git
cd RTUPointListParser
```

### 2. Set Up Tessdata

Ensure the `tessdata` directory exists in the project root with the English trained data file:

```
RTUPointListParser/
├── tessdata/
│   └── eng.traineddata
├── RTUPointlistParse/
└── ...
```

You can download `eng.traineddata` from: https://github.com/tesseract-ocr/tessdata

### 3. Build the Project

```bash
cd RTUPointlistParse/RTUPointlistParse
dotnet build
```

## Usage

### Command Line

```bash
dotnet run --project RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj [inputFolder] [outputFolder]
```

### Parameters

- `inputFolder` (optional): Path to folder containing PDF files
  - Default: `C:\dev\RTUPointListParser\ExamplePointlists\Example1\Input`
- `outputFolder` (optional): Path to output folder
  - Default: `C:\dev\RTUPointListParser\ExamplePointlists\Example1\TestOutput`

### Environment Variables

- `TESSDATA_DIR`: Path to tessdata directory (default: project root tessdata folder)
- `TESSERACT_LANG`: Tesseract language (default: "eng")

### Example

**Using default folders:**
```bash
dotnet run
```

**Using custom folders:**
```bash
dotnet run "C:\path\to\pdfs" "C:\path\to\output"
```

## Output Files

The application generates two files in the output folder:

### 1. PointList.xlsx
Excel file with a single sheet named "Points" containing:
- **Point Number**: Integer point number (sorted ascending)
- **Point Name**: String description of the point

### 2. PointList.log.txt
Processing log with:
- Files processed
- Rows extracted
- Validation warnings (duplicates, gaps in sequence)
- Any errors encountered

## How It Works

### 1. PDF Rendering
Uses **PdfiumViewer** to render PDF pages as high-resolution bitmaps (~300 DPI).

### 2. OCR Processing
Uses **Tesseract** to extract text with word-level bounding boxes, capturing both text content and spatial positioning.

### 3. Table Detection
Analyzes word positions to identify column bands:
- **Column 1**: Point Number (leftmost numeric column, may have stacked "POINT"/"NUMBER" header)
- **Column 2**: Point Name (next column to the right)

### 4. Row Clustering
Groups words into rows based on vertical (Y) proximity, then extracts data from each row within the detected column bands.

### 5. Data Extraction
For each row:
- Extracts the first numeric token from Column 1 as Point Number
- Combines all words from Column 2 (ordered by X position) as Point Name
- Filters out malformed rows and "SPARE" entries

### 6. Validation & Output
- Aggregates rows from all PDFs and pages
- Sorts by Point Number (ascending)
- Validates continuity (logs gaps and duplicates)
- Writes Excel file and log

## Implementation Details

### Key Classes

#### `App` Class
Main application class with the following methods:

- `Task<int> RunAsync(...)`: Main processing pipeline
- `GetPdfFiles(string)`: Enumerates PDF files in input folder
- `RenderPdfToBitmaps(string, int, Action<string>)`: Converts PDF pages to bitmaps
- `OcrWordsFromBitmap(Bitmap, TesseractEngine)`: Performs OCR and extracts word bounding boxes
- `DetectFirstTwoColumnBands(List<OcrWord>)`: Identifies column regions by geometry
- `ClusterWordsIntoRows(List<OcrWord>, int)`: Groups words into row clusters
- `ExtractRows(...)`: Extracts Point Number and Point Name from rows
- `Normalize(string)`: Cleans up OCR text artifacts
- `ValidateSequence(IEnumerable<int>)`: Detects gaps and duplicates
- `WriteExcel(string, IEnumerable<(int, string)>)`: Generates Excel output
- `WriteLog(string, IEnumerable<string>)`: Writes log file

#### Records
- `OcrWord(string Text, Rectangle Bounds, float Confidence)`: Represents a word from OCR with spatial info
- `RowCluster(int Y, List<OcrWord> Words)`: Represents a horizontal row of words

### Dependencies

```xml
<PackageReference Include="PdfiumViewer" Version="2.13.0" />
<PackageReference Include="Tesseract" Version="5.2.0" />
<PackageReference Include="ClosedXML" Version="0.105.0" />
<PackageReference Include="System.Drawing.Common" Version="8.0.0" />
```

## Project Structure

```
RTUPointListParser/
├── tessdata/
│   └── eng.traineddata          # Tesseract language data
├── RTUPointlistParse/
│   └── RTUPointlistParse/
│       ├── Program.cs            # Main application code
│       └── RTUPointlistParse.csproj
├── ExamplePointlists/
│   └── Example1/
│       ├── Input/                # Sample PDF files
│       ├── TestOutput/           # Generated output
│       └── Expected Output/      # Reference output
└── README.md
```

## Troubleshooting

### Build Warnings
- **NU1701**: PdfiumViewer targets .NET Framework but is compatible with .NET 8 - this warning is safe to ignore.

### Runtime Issues
- **Native Library Errors**: PdfiumViewer requires Windows native libraries. This application cannot run on Linux or macOS.
- **Tesseract Not Found**: Ensure `TESSDATA_DIR` environment variable points to a valid tessdata directory.
- **Poor OCR Results**: 
  - Verify input PDFs are scanned at adequate resolution (300 DPI recommended)
  - Check that `eng.traineddata` is the correct version for Tesseract 5.x

### Empty Output
- Verify PDFs contain image data (not just vector graphics)
- Check log file for error messages
- Ensure table structure matches expected format (numeric column followed by text column)

## Example Output

### Console Output
```
RTU Point List Parser
=====================
Input folder: C:\dev\RTUPointListParser\ExamplePointlists\Example1\Input
Output folder: C:\dev\RTUPointListParser\ExamplePointlists\Example1\TestOutput

Found 2 PDF file(s) to process

Processing: Control115_sh1_145934779.pdf
  Rendered 1 page(s)
  Extracted 97 total rows so far
Processing: Control115_sh2_145937196.pdf
  Rendered 1 page(s)
  Extracted 197 total rows so far

Total rows extracted: 197

Written: C:\dev\RTUPointListParser\ExamplePointlists\Example1\TestOutput\PointList.xlsx
Written: C:\dev\RTUPointListParser\ExamplePointlists\Example1\TestOutput\PointList.log.txt
```

## License

This project is provided as-is for educational and internal use.

## Contributing

Contributions are welcome! Please ensure all changes maintain the existing code structure and include appropriate tests.
