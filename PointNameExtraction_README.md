# Point Name Extraction

This document explains how to extract "Point Name" column values from PDF files in the RTUPointListParser project.

## Overview

The `extract_point_names.py` script extracts all values under the "Point Name" column header from every table in each PDF file. It:

- Processes all PDF files in the input folder
- Extracts text using OCR (Optical Character Recognition) for image-based PDFs
- Identifies and extracts "Point Name" column values from table structures
- Filters out empty rows and rows where "Point Name" is "Spare"
- Outputs results to an Excel file with Status and Analog sheets

## Prerequisites

### System Dependencies

The script requires OCR tools to be installed:

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr poppler-utils
```

**macOS:**
```bash
brew install tesseract poppler
```

**Windows:**
1. Install Tesseract OCR from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install Poppler Utils from: https://blog.alivate.com.au/poppler-windows/
3. Add both to your system PATH

### Python Dependencies

```bash
pip install PyPDF2 pdfplumber openpyxl
```

## Usage

### Basic Usage

```bash
python3 extract_point_names.py
```

This uses default folders:
- Input: `ExamplePointlists/Example1/Input`
- Output: `ExamplePointlists/Example1/TestOutput`

### Custom Folders

```bash
python3 extract_point_names.py <input_folder> <output_folder>
```

Example:
```bash
python3 extract_point_names.py path/to/pdfs path/to/output
```

## Output

The script generates an Excel file named `Point_Names_Extracted.xlsx` with two worksheets:

- **Status**: Point names from Status sheet PDFs (typically files with "sh1" or "status" in filename)
- **Analog**: Point names from Analog sheet PDFs (typically files with "sh2" or "analog" in filename)

Each worksheet contains:
- Title row
- Description row
- Column header ("POINT NAME")
- Extracted point name values (one per row)

## Filtering Rules

The extraction process:

1. **Ignores empty rows**: Rows with no point name value are skipped
2. **Filters "Spare" entries**: Rows where the point name is "Spare" (case-insensitive) are excluded
3. **Validates point names**: Filters out OCR artifacts and invalid entries

## How It Works

1. **PDF Processing**: For each PDF file:
   - Attempts direct text extraction using pdfplumber
   - Falls back to PyPDF2 if pdfplumber fails
   - Uses OCR (tesseract + pdftoppm) for image-based PDFs

2. **Text Analysis**: 
   - Identifies table structures by looking for "POINT NAME" headers
   - Parses data rows (lines starting with numbers)
   - Extracts point name values from the appropriate column

3. **Sheet Classification**:
   - Files with "sh1" or "status" in name → Status sheet
   - Files with "sh2" or "analog" in name → Analog sheet
   - Others default to Status sheet

4. **Output Generation**:
   - Creates Excel workbook with Status and Analog sheets
   - Formats with headers and titles
   - Saves to output folder

## Limitations

- OCR quality depends on PDF scan quality and resolution
- Some OCR artifacts may remain in extracted text
- Table structure recognition is based on common patterns and may not work for all PDF formats
- Multiple tables in a single PDF are processed sequentially (all point names from all tables are combined)

## Example

```bash
$ python3 extract_point_names.py ExamplePointlists/Example1/Input ExamplePointlists/Example1/TestOutput

Point Name Extractor
====================
Input folder: ExamplePointlists/Example1/Input
Output folder: ExamplePointlists/Example1/TestOutput

Found 2 PDF file(s) to process

Processing: Control115_sh1_145934779.pdf
  Using OCR to extract text...
  OCR completed on 1 page(s)
  Extracted 53 point names (Status)
Processing: Control115_sh2_145937196.pdf
  Using OCR to extract text...
  OCR completed on 1 page(s)
  Extracted 50 point names (Analog)

Creating output Excel file...
  Status: 53 point names
  Analog: 50 point names
Output saved to: ExamplePointlists/Example1/TestOutput/Point_Names_Extracted.xlsx

Extraction complete.
```

## Troubleshooting

### "tesseract not found" or "pdftoppm not found"

Ensure OCR tools are installed and in your system PATH. After installation, restart your terminal/IDE.

Verify installation:
```bash
tesseract --version
pdftoppm -v
```

### Poor extraction quality

OCR quality depends on:
- PDF scan resolution (higher is better)
- Image clarity and contrast
- Font size and style

For better results, ensure source PDFs are high-quality scans.

### Missing point names

If point names are missing, the extraction logic may not recognize the table structure. Check:
- Does the PDF contain text under a "POINT NAME" header?
- Are rows structured with numbers at the start?
- Are there OCR errors preventing pattern matching?

## Integration with Existing C# Application

This Python script is a standalone tool focused specifically on extracting Point Name values. The main C# application (`RTUPointlistParse`) provides full table parsing and Excel generation with all columns.

Use the Python script when you only need Point Name extraction, or as a quick validation tool.
