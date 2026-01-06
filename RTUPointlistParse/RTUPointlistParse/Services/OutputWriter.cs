using RTUPointlistParse.Models;
using System.Text;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for generating output from point records.
/// </summary>
public class OutputWriter
{
    /// <summary>
    /// Generates formatted output from a list of point records.
    /// </summary>
    /// <param name="records">List of point records to format.</param>
    /// <returns>Formatted output string.</returns>
    public string GenerateOutput(List<PointRecord> records)
    {
        // Sort records for deterministic output
        var sortedRecords = records
            .OrderBy(r => r.Name)
            .ThenBy(r => r.Address)
            .ToList();

        var output = new StringBuilder();

        // Write header
        output.AppendLine("Name,Address,Type,Unit,Description");

        // Write records in CSV format
        foreach (var record in sortedRecords)
        {
            output.AppendLine($"{EscapeCsv(record.Name)},{EscapeCsv(record.Address)},{EscapeCsv(record.Type)},{EscapeCsv(record.Unit)},{EscapeCsv(record.Description)}");
        }

        return output.ToString();
    }

    /// <summary>
    /// Writes output to a file.
    /// </summary>
    /// <param name="content">Content to write.</param>
    /// <param name="filePath">Output file path.</param>
    public void WriteToFile(string content, string filePath)
    {
        // Use UTF-8 encoding without BOM
        var utf8WithoutBom = new UTF8Encoding(false);
        File.WriteAllText(filePath, content, utf8WithoutBom);
    }

    /// <summary>
    /// Escapes a CSV field value.
    /// </summary>
    private string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        // If value contains comma, newline, or quote, wrap in quotes and escape internal quotes
        if (value.Contains(',') || value.Contains('\n') || value.Contains('\r') || value.Contains('"'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return value;
    }
}
