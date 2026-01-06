# RTU Point List Parser

A C# .NET Console Application for parsing RTU point list data from PDF files and generating Excel (.xlsx) output files.

## Features

- **PDF Text Extraction**: Extracts text from PDF files using PdfPig library
- **Table Parsing**: Converts extracted text into structured table data
- **Excel Generation**: Creates formatted .xlsx files with Status and Analog sheets using ClosedXML
- **File Comparison**: Compares generated Excel files against expected output
- **Batch Processing**: Processes all PDF files in the input folder and combines them into a single output

## Usage

### Command Line

```bash
dotnet run --project RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj [inputFolder] [outputFolder]
```

### Parameters

- `inputFolder` (optional): Path to folder containing PDF files to parse
  - Default: `ExamplePointlists/Example1/Input`
- `outputFolder` (optional): Path to folder where Excel files will be saved
  - Default: `ExamplePointlists/Example1/TestOutput`

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

The application generates Excel files with two worksheets:

### Status Sheet
Contains DNP status point list data with columns including:
- TAB DEC DNP INDEX
- CONTROL ADDRESS
- POINT NAME
- NORMAL STATE
- STATE INFO
- ALARMS
- EMS CROSS REFERENCE DATA
- IED INFORMATION

### Analog Sheet
Contains DNP analog point list data with columns including:
- TAB DEC DNP INDEX
- POINT NAME
- SCALING (Coefficient, Offset)
- FULL SCALE (Value, Unit)
- LIMITS (Low, High)
- ALARMS
- EMS CROSS REFERENCE DATA
- IED INFORMATION

## Helper Methods

The application implements the following key helper methods:

### `ExtractTextFromPdf(string filePath)`
Extracts text content from a PDF file using PdfPig library.

**Parameters:**
- `filePath`: Full path to the PDF file

**Returns:** String containing extracted text

### `ParseTable(string pdfText)`
Parses table data from extracted PDF text into structured rows.

**Parameters:**
- `pdfText`: Text extracted from PDF

**Returns:** List of `TableRow` objects containing column data

### `GenerateExcel(List<TableRow> statusRows, List<TableRow> analogRows, string outputPath)`
Creates an Excel (.xlsx) file with formatted Status and Analog sheets.

**Parameters:**
- `statusRows`: Data rows for Status sheet
- `analogRows`: Data rows for Analog sheet
- `outputPath`: Full path where Excel file will be saved

### `CompareExcelFiles(string generatedFile, string expectedFile)`
Compares two Excel files and reports differences.

**Parameters:**
- `generatedFile`: Path to generated Excel file
- `expectedFile`: Path to expected Excel file

**Returns:** Boolean indicating if files match

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

```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr poppler-utils

# macOS
brew install tesseract poppler

# Windows
# Install from: https://github.com/UB-Mannheim/tesseract/wiki
# Install from: https://blog.alivate.com.au/poppler-windows/
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

1. Process all PDF files in `ExamplePointlists/Example1/Input`
2. Generate `Control_rtu837_DNP_pointlist_rev00.xlsx` in `ExamplePointlists/Example1/TestOutput`
3. Compare the generated file with the expected output in `ExamplePointlists/Example1/Expected Output`
4. Display a summary of any differences found

## Building

```bash
dotnet build RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj
```

## Requirements

- .NET 10.0 or later
- Linux, macOS, or Windows
