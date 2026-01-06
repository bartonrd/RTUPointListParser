using RTUPointlistParse.Models;
using System.Text;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for generating output files from point records.
/// </summary>
public class OutputWriter
{
    /// <summary>
    /// Generates output content from a list of point records.
    /// </summary>
    /// <param name="records">List of point records to format.</param>
    /// <param name="format">Output format (csv, json, or txt).</param>
    /// <returns>Formatted output string.</returns>
    public string GenerateOutput(List<PointRecord> records, string format = "csv")
    {
        if (records == null || records.Count == 0)
        {
            return "No records to output.";
        }

        // Sort records for deterministic output
        var sortedRecords = records
            .OrderBy(r => r.Name)
            .ThenBy(r => r.Address)
            .ToList();

        return format.ToLowerInvariant() switch
        {
            "csv" => GenerateCsv(sortedRecords),
            "json" => GenerateJson(sortedRecords),
            "txt" => GenerateText(sortedRecords),
            _ => GenerateCsv(sortedRecords)
        };
    }

    /// <summary>
    /// Writes output to a file.
    /// </summary>
    /// <param name="content">Content to write.</param>
    /// <param name="filePath">Path to output file.</param>
    public void WriteToFile(string content, string filePath)
    {
        try
        {
            // Ensure output directory exists
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Write with UTF-8 encoding without BOM
            var utf8WithoutBom = new UTF8Encoding(false);
            File.WriteAllText(filePath, content, utf8WithoutBom);
            
            Console.WriteLine($"Successfully wrote output to: {filePath}");
        }
        catch (Exception ex)
        {
            throw new IOException($"Failed to write output to '{filePath}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Generates CSV format output.
    /// </summary>
    private string GenerateCsv(List<PointRecord> records)
    {
        var sb = new StringBuilder();

        // Header
        sb.AppendLine("Name,Address,Type,Unit,Description,NormalState,State1,State0,Alarm,EmsNumber");

        // Data rows
        foreach (var record in records)
        {
            sb.AppendLine(string.Join(",",
                EscapeCsvValue(record.Name),
                EscapeCsvValue(record.Address),
                EscapeCsvValue(record.Type),
                EscapeCsvValue(record.Unit),
                EscapeCsvValue(record.Description),
                EscapeCsvValue(record.NormalState),
                EscapeCsvValue(record.State1),
                EscapeCsvValue(record.State0),
                EscapeCsvValue(record.Alarm),
                EscapeCsvValue(record.EmsNumber)
            ));
        }

        return sb.ToString();
    }

    /// <summary>
    /// Generates JSON format output.
    /// </summary>
    private string GenerateJson(List<PointRecord> records)
    {
        var sb = new StringBuilder();
        sb.AppendLine("[");

        for (int i = 0; i < records.Count; i++)
        {
            var record = records[i];
            sb.AppendLine("  {");
            sb.AppendLine($"    \"Name\": {EscapeJsonValue(record.Name)},");
            sb.AppendLine($"    \"Address\": {EscapeJsonValue(record.Address)},");
            sb.AppendLine($"    \"Type\": {EscapeJsonValue(record.Type)},");
            sb.AppendLine($"    \"Unit\": {EscapeJsonValue(record.Unit)},");
            sb.AppendLine($"    \"Description\": {EscapeJsonValue(record.Description)},");
            sb.AppendLine($"    \"NormalState\": {EscapeJsonValue(record.NormalState)},");
            sb.AppendLine($"    \"State1\": {EscapeJsonValue(record.State1)},");
            sb.AppendLine($"    \"State0\": {EscapeJsonValue(record.State0)},");
            sb.AppendLine($"    \"Alarm\": {EscapeJsonValue(record.Alarm)},");
            sb.AppendLine($"    \"EmsNumber\": {EscapeJsonValue(record.EmsNumber)}");
            sb.Append("  }");
            if (i < records.Count - 1)
                sb.AppendLine(",");
            else
                sb.AppendLine();
        }

        sb.AppendLine("]");
        return sb.ToString();
    }

    /// <summary>
    /// Generates plain text format output.
    /// </summary>
    private string GenerateText(List<PointRecord> records)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Point List Output");
        sb.AppendLine("=================");
        sb.AppendLine();

        foreach (var record in records)
        {
            sb.AppendLine($"Name: {record.Name}");
            sb.AppendLine($"Address: {record.Address}");
            sb.AppendLine($"Type: {record.Type}");
            sb.AppendLine($"Unit: {record.Unit}");
            sb.AppendLine($"Description: {record.Description}");
            sb.AppendLine($"Normal State: {record.NormalState}");
            sb.AppendLine($"State 1: {record.State1}");
            sb.AppendLine($"State 0: {record.State0}");
            sb.AppendLine($"Alarm: {record.Alarm}");
            sb.AppendLine($"EMS Number: {record.EmsNumber}");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// Escapes a value for CSV format.
    /// </summary>
    private string EscapeCsvValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "\"\"";

        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return $"\"{value}\"";
    }

    /// <summary>
    /// Escapes a value for JSON format.
    /// </summary>
    private string EscapeJsonValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "\"\"";

        var escaped = value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");

        return $"\"{escaped}\"";
    }
}
