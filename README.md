# RTUPointListParser

A C# console application that processes RTU point list PDF files and generates structured Excel output.

## Repository Structure

```
RTUPointListParser/
├── PointListProcessor/          # C# Console Application
│   ├── Program.cs               # Main application logic
│   ├── PointListProcessor.csproj # Project file
│   └── README.md                # Detailed application documentation
├── ExamplePointlists/           # Test examples
│   └── Example1/
│       ├── Input/               # Sample PDF files
│       ├── Expected Output/     # Reference output files
│       └── TestOutput/          # Generated output files
└── PointListProcessor.sln       # Visual Studio solution file
```

## Quick Start

1. **Build the application:**
   ```bash
   cd PointListProcessor
   dotnet build
   ```

2. **Run with example files:**
   ```bash
   dotnet run -- --input "../ExamplePointlists/Example1/Input" --output "../ExamplePointlists/Example1/TestOutput"
   ```

3. **Or run in interactive mode:**
   ```bash
   dotnet run
   ```

## Features

- ✅ Command-line argument support (`--input` and `--output` flags)
- ✅ Interactive prompts for user input
- ✅ Default directories (`Input` and `TestOutput`)
- ✅ PDF text extraction using iText7
- ✅ Excel output generation using EPPlus
- ✅ Error handling and validation
- ✅ Processes multiple PDF files in batch

## Usage Examples

### Using command-line flags:
```bash
PointListProcessor.exe --input "C:\MyInput" --output "C:\MyOutput"
```

### Using positional arguments:
```bash
PointListProcessor.exe "C:\MyInput" "C:\MyOutput"
```

### Using interactive mode:
```bash
PointListProcessor.exe
# Then press Enter twice to use defaults, or provide custom paths
```

## Documentation

For detailed documentation, see [PointListProcessor/README.md](PointListProcessor/README.md).

## License

This project uses EPPlus under the Non-Commercial License.
