using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for extracting text from PDF files.
/// </summary>
public class PdfTextExtractor
{
    /// <summary>
    /// Extracts text from a PDF file.
    /// </summary>
    /// <param name="filePath">Path to the PDF file.</param>
    /// <returns>Extracted text content.</returns>
    /// <exception cref="NotSupportedException">Thrown when PDF contains only images (OCR required).</exception>
    public string ExtractTextFromPdf(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"PDF file not found: {filePath}");
        }

        try
        {
            using var document = PdfDocument.Open(filePath);
            var allText = new System.Text.StringBuilder();

            foreach (var page in document.GetPages())
            {
                var text = page.Text;
                
                // Check if page has extractable text
                if (string.IsNullOrWhiteSpace(text))
                {
                    var words = page.GetWords();
                    if (!words.Any())
                    {
                        throw new NotSupportedException(
                            $"PDF file '{Path.GetFileName(filePath)}' appears to be image-based. " +
                            "OCR (Optical Character Recognition) is required to extract text from scanned documents. " +
                            "This feature is not yet implemented. " +
                            "Consider converting the PDF to text or using an OCR tool first.");
                    }
                }

                allText.AppendLine(text);
            }

            return allText.ToString();
        }
        catch (NotSupportedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to extract text from PDF '{Path.GetFileName(filePath)}': {ex.Message}", 
                ex);
        }
    }

    /// <summary>
    /// Checks if a PDF file contains extractable text.
    /// </summary>
    /// <param name="filePath">Path to the PDF file.</param>
    /// <returns>True if text can be extracted, false if OCR is needed.</returns>
    public bool HasExtractableText(string filePath)
    {
        try
        {
            using var document = PdfDocument.Open(filePath);
            
            foreach (var page in document.GetPages())
            {
                var words = page.GetWords();
                if (words.Any())
                {
                    return true;
                }
            }
            
            return false;
        }
        catch
        {
            return false;
        }
    }
}
