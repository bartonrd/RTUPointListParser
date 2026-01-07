# RTU Point List Parser

A C# .NET Console Application for parsing RTU point list data from **image-based PDF files** and generating Excel (.xlsx) output files with sequential point numbering.

## Features

- **Image-Based PDF Processing**: Handles scanned/image-based PDF files using OCR technology
- **Multi-Page Support**: Processes PDFs with any number of pages
- **Multi-Table Support**: Extracts data from all tables in all PDFs
- **PDF Text Extraction**: Extracts text from PDF files using PdfPig library with OCR fallback
- **Sequential Point Numbering**: Automatically sorts and renumbers points from 1 to N
- **Table Parsing**: Converts extracted text into structured table data
- **Excel Generation**: Creates formatted .xlsx files with Status and Analog sheets using ClosedXML
- **Batch Processing**: Processes all PDF files in the input folder and combines them into a single output

## Usage

### Command Line

```bash
dotnet run --project RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj [inputFolder] [outputFolder]
```

### Parameters

- `inputFolder` (optional): Path to folder containing PDF files to parse
  - Default: `C:\dev\RTUPointListParser\ExamplePointlists\Example1\Input`
- `outputFolder` (optional): Path to folder where Excel files will be saved
  - Default: `C:\dev\RTUPointListParser\ExamplePointlists\Example1\TestOutput`

### Examples

**Using default folders:**
```bash
dotnet run --project RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj
```

**Using custom folders:**
```bash
dotnet run --project RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj "path/to/input" "path/to/output"
```

## Output Format

The application generates a single Excel file containing data from all processed PDFs, with two worksheets:

### Status Sheet
Contains DNP status point list data with columns:
- **Point Number**: Sequential number from 1 to N
- **Point Name**: Name/description of the status point

### Analog Sheet
Contains DNP analog point list data with columns:
- **Point Number**: Sequential number from 1 to N
- **Point Name**: Name/description of the analog point

**Note**: 
- The parser automatically extracts Point Number from the PDF files
- Points are sorted by their original number and then renumbered sequentially starting from 1
- All points are included (including Spare entries)
- The Point Number column in PDFs has "Point" stacked above "Number" in the header

## Helper Methods

The application implements the following key helper methods:

### `ExtractTextFromPdf(string filePath)`
Extracts text content from a PDF file using PdfPig library. If no text is found (image-based PDF), automatically falls back to OCR.

**Parameters:**
- `filePath`: Full path to the PDF file

**Returns:** String containing extracted text from all pages

### `ParseStatusTable(string pdfText)` and `ParseAnalogTable(string pdfText)`
Parses table data from extracted PDF text into structured rows, extracting Point Number and Point Name columns.

**Parameters:**
- `pdfText`: Text extracted from PDF

**Returns:** List of `TableRow` objects containing Point Number and Point Name

### `SortAndRenumberRows(List<TableRow> rows)`
Sorts rows by their extracted point number and renumbers them sequentially starting from 1.

**Parameters:**
- `rows`: List of table rows to sort and renumber

**Returns:** Sorted and renumbered list of rows

### `GenerateExcel(List<TableRow> statusRows, List<TableRow> analogRows, string outputPath)`
Creates an Excel (.xlsx) file with Status and Analog sheets, each containing Point Number and Point Name columns.

**Parameters:**
- `statusRows`: Data rows for Status sheet (Point Number and Point Name)
- `analogRows`: Data rows for Analog sheet (Point Number and Point Name)
- `outputPath`: Full path where Excel file will be saved

## Libraries Used

- **PdfPig (1.7.0-custom-5)**: PDF text extraction
- **ClosedXML (0.105.0)**: Excel file generation and manipulation
- **Tesseract OCR**: Optical Character Recognition for image-based PDFs (via system binary)
- **Poppler Utils**: PDF to image conversion (via pdftoppm)

## OCR Support for Image-Based PDFs

The application automatically handles image-based PDFs (scanned documents) using OCR:

### How It Works

1. **Text Extraction First**: Attempts to extract text directly from the PDF using PdfPig
2. **OCR Fallback**: If no text is found, automatically:
   - Converts PDF pages to images using `pdftoppm`
   - Performs OCR using Tesseract
   - Extracts text from the images
3. **Data Parsing**: Parses the extracted text into structured table data

### System Requirements for OCR

The application requires the following system packages for OCR functionality:

#### Windows

1. **Tesseract OCR**:
   - Download installer from: https://github.com/UB-Mannheim/tesseract/wiki
   - Run the installer (recommended path: `C:\Program Files\Tesseract-OCR`)
   - **Important**: During installation, check "Add to PATH" or manually add to system PATH
   - Verify installation: Open Command Prompt and run `tesseract --version`

2. **Poppler Utils**:
   - Download from: https://blog.alivate.com.au/poppler-windows/
   - Extract the ZIP file (e.g., to `C:\Program Files\poppler`)
   - Add the `bin` folder to your system PATH (e.g., `C:\Program Files\poppler\bin`)
   - Verify installation: Open Command Prompt and run `pdftoppm -v`

**Adding to PATH on Windows:**
- Right-click "This PC" → Properties → Advanced system settings → Environment Variables
- Under "System variables", find "Path" and click Edit
- Click "New" and add the path to the tool's bin folder
- Click OK and restart your Command Prompt/PowerShell

#### Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr poppler-utils
```

#### macOS

```bash
brew install tesseract poppler
```

### OCR Performance

- **Accuracy**: Depends on PDF image quality and resolution
- **Speed**: Processes approximately 1-2 pages per second
- **Output**: Extracts text data even from scanned/image-only PDFs

### Example OCR Output

```
Processing: Control115_sh2_145937196.pdf
  No text found, attempting OCR...
  Performing OCR on PDF...
  OCR completed on 1 page(s)
  Extracted 100 Analog rows

Processing: Control115_sh1_145934779.pdf
  No text found, attempting OCR...
  Performing OCR on PDF...
  OCR completed on 1 page(s)
  Extracted 97 Status rows
```

## Example Output

When run with the Example1 data, the application will:

1. Process all PDF files in the input folder (e.g., `Control115_sh1_145934779.pdf`, `Control115_sh2_145937196.pdf`)
2. Use OCR to extract text from image-based PDFs
3. Parse Point Number and Point Name from all tables in all pages
4. Sort points by their extracted number and renumber sequentially from 1 to N
5. Generate `Control_rtu837_DNP_pointlist_rev00.xlsx` in the output folder with:
   - **Status Sheet**: Sequential points with Point Number (1-N) and Point Name
   - **Analog Sheet**: Sequential points with Point Number (1-N) and Point Name

Sample console output:
```
RTU Point List Parser
=====================
Input folder: ExamplePointlists/Example1/Input
Output folder: ExamplePointlists/Example1/TestOutput

Found 2 PDF file(s) to process

Processing: Control115_sh2_145937196.pdf
  No text found, attempting OCR...
  Performing OCR on PDF...
  OCR completed on 1 page(s)
  Extracted 40 Analog rows
Processing: Control115_sh1_145934779.pdf
  No text found, attempting OCR...
  Performing OCR on PDF...
  OCR completed on 1 page(s)
  Extracted 72 Status rows

Generating combined Excel file: Control_rtu837_DNP_pointlist_rev00.xlsx
  Status points: 72
  Analog points: 40
  Generated: Control_rtu837_DNP_pointlist_rev00.xlsx

Processing complete.
```

## Building

```bash
dotnet build RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj
```

## Requirements

- .NET 10.0 or later
- Linux, macOS, or Windows
