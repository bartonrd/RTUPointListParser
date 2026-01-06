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

## Known Limitations

### Image-Based PDFs
The application uses text extraction from PDFs. If the PDF files are image-based (scanned documents or graphics-only), the text extraction will not work. In such cases:

- The application will display a warning: "No text extracted from PDF. The PDF may be image-based."
- An Excel file will still be generated with the proper template structure
- The Excel file will contain headers but no data rows

**Workaround for image-based PDFs:**
To process image-based PDFs, you would need to implement OCR (Optical Character Recognition) using additional libraries such as:
- Tesseract.NET
- Azure Computer Vision API
- Google Cloud Vision API

This is beyond the scope of the current implementation.

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
