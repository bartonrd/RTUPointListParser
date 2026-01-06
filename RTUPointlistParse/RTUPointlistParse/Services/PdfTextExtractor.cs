using UglyToad.PdfPig;
using System.Text;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for extracting text from PDF files.
/// </summary>
public class PdfTextExtractor
{
    /// <summary>
    /// Extracts text content from a PDF file.
    /// </summary>
    /// <param name="filePath">Path to the PDF file.</param>
    /// <returns>Extracted text content.</returns>
    /// <exception cref="NotSupportedException">Thrown when file type is not supported.</exception>
    public string ExtractTextFromPdf(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLower();

        // Handle text files directly
        if (extension == ".txt" || extension == ".csv")
        {
            return File.ReadAllText(filePath);
        }

        // Handle PDF files using PdfPig
        if (extension == ".pdf")
        {
            return ExtractTextFromPdfUsingPdfPig(filePath);
        }

        throw new NotSupportedException($"File type '{extension}' is not supported: {filePath}");
    }

    /// <summary>
    /// Extracts text from a PDF file using PdfPig library.
    /// </summary>
    private string ExtractTextFromPdfUsingPdfPig(string filePath)
    {
        var text = new StringBuilder();

        using (var document = PdfDocument.Open(filePath))
        {
            foreach (var page in document.GetPages())
            {
                var pageText = page.Text;
                text.AppendLine(pageText);
            }
        }

        return text.ToString();
    }
}
