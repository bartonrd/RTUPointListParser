# RTU Point List Parser

A C# .NET Console Application for parsing RTU point list data from PDF files and generating Excel (.xlsx) output files.

## Features

- **PDF Text Extraction**: Extracts text from PDF files using PdfPig library
- **OCR Support**: Automatically handles image-based PDFs using Tesseract OCR
- **Point Name Extraction**: Extracts Point Names from all tables in PDF files
- **Smart Filtering**: Automatically filters out "Spare" entries and empty rows
- **OCR Artifact Cleanup**: Cleans up common OCR recognition errors
- **Excel Generation**: Creates formatted .xlsx files with extracted data using ClosedXML
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

The application generates an Excel file with a single worksheet containing extracted Point Names:

### Point Names Sheet
Contains a single column with all Point Names extracted from the PDF files:
- Point Name (header)
- One point name per row
- "Spare" entries are filtered out
- Empty rows are excluded
- OCR artifacts are cleaned up automatically

## Helper Methods

The application implements the following key helper methods:

### `ExtractTextFromPdf(string filePath)`
Extracts text content from a PDF file using PdfPig library. Automatically falls back to OCR if no text is found.

**Parameters:**
- `filePath`: Full path to the PDF file

**Returns:** String containing extracted text

### `ExtractPointNamesFromPdf(string pdfText)`
Extracts only Point Names from PDF text, filtering out "Spare" entries and empty rows.

**Parameters:**
- `pdfText`: Text extracted from PDF

**Returns:** List of strings containing point names

### `GeneratePointNameExcel(List<string> pointNames, string outputPath)`
Creates an Excel (.xlsx) file with a single column containing Point Names.

**Parameters:**
- `pointNames`: List of point names to include in the Excel file
- `outputPath`: Full path where Excel file will be saved

### `FinalCleanPointName(string pointName)`
Performs final cleanup on extracted point names to remove OCR artifacts.

**Parameters:**
- `pointName`: Raw point name with potential OCR errors

**Returns:** Cleaned point name string
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
3. **Point Name Extraction**: Parses the extracted text to identify Point Name columns
4. **OCR Cleanup**: Applies intelligent cleanup to remove common OCR artifacts
   - Removes leading lowercase letters before uppercase (e.g., "l115KV" → "115KV")
   - Removes erroneous leading capital letters (e.g., "INO." → "NO.")
   - Fixes common character confusions (e.g., "N15KV" → "115KV")
   - Filters out placeholder entries and formatting artifacts

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
  Extracted 34 Point Names

Processing: Control115_sh1_145934779.pdf
  No text found, attempting OCR...
  Performing OCR on PDF...
  OCR completed on 1 page(s)
  Extracted 41 Point Names

Generating Excel file with Point Names: Control_rtu837_DNP_pointlist_rev00.xlsx
Total Point Names extracted: 75
  Generated: Control_rtu837_DNP_pointlist_rev00.xlsx
```

## Example Output

When run with the Example1 data, the application will:

1. Process all PDF files in `ExamplePointlists/Example1/Input`
2. Extract Point Names from all tables in each PDF
3. Filter out "Spare" entries and empty rows
4. Apply OCR artifact cleanup
5. Generate `Control_rtu837_DNP_pointlist_rev00.xlsx` in `ExamplePointlists/Example1/TestOutput`
   - Single worksheet named "Point Names"
   - One column with header "Point Name"
   - One point name per row (75 point names from the example PDFs)

## Building

```bash
dotnet build RTUPointlistParse/RTUPointlistParse/RTUPointlistParse.csproj
```

## Requirements

- .NET 10.0 or later
- Linux, macOS, or Windows
- For OCR support (image-based PDFs):
  - Tesseract OCR
  - Poppler Utils (pdftoppm)
