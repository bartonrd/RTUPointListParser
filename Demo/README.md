# Demo Folder

This folder contains a demonstration of the PointListProcessor application with sample data.

## Structure

- `Input/`: Contains sample input files for testing
- `Output/`: Contains generated output files (created by the application)

## Usage

To test the application with this demo:

```bash
cd PointListProcessor/bin/Debug/net8.0
dotnet PointListProcessor.dll --input ../../../../../Demo/Input --output ../../../../../Demo/Output
```

Or from the project root:

```bash
cd PointListProcessor
dotnet run -- --input ../Demo/Input --output ../Demo/Output
```

## Note

The sample PDFs in the repository (ExamplePointlists/Example1/Input/) appear to be scanned images without text layers, so they won't extract text content. To properly test the application, you'll need PDF files with extractable text or use text-based point lists.
