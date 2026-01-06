using RTUPointlistParse.Models;
using System.Text.RegularExpressions;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for parsing point list data into structured records.
/// </summary>
public class PointListParser
{
    /// <summary>
    /// Parses text content into a list of point records.
    /// </summary>
    /// <param name="content">Text content to parse.</param>
    /// <returns>List of parsed point records.</returns>
    public List<PointRecord> ParsePointList(string content)
    {
        var records = new List<PointRecord>();
        
        if (string.IsNullOrWhiteSpace(content))
        {
            Console.WriteLine("Warning: Empty content provided for parsing.");
            return records;
        }

        var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        int lineNumber = 0;
        int parsedCount = 0;
        int skippedCount = 0;

        foreach (var line in lines)
        {
            lineNumber++;
            var trimmedLine = line.Trim();

            // Skip empty lines, headers, and common non-data lines
            if (string.IsNullOrWhiteSpace(trimmedLine) ||
                IsHeaderOrFooter(trimmedLine) ||
                trimmedLine.Length < 10)
            {
                continue;
            }

            try
            {
                var record = ParseLine(trimmedLine);
                if (record != null)
                {
                    records.Add(record);
                    parsedCount++;
                }
                else
                {
                    skippedCount++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to parse line {lineNumber}: {ex.Message}");
                Console.WriteLine($"  Line content: {trimmedLine.Substring(0, Math.Min(50, trimmedLine.Length))}...");
                skippedCount++;
            }
        }

        Console.WriteLine($"Parsing complete: {parsedCount} records parsed, {skippedCount} lines skipped.");
        return records;
    }

    /// <summary>
    /// Determines if a line is a header or footer that should be skipped.
    /// </summary>
    private bool IsHeaderOrFooter(string line)
    {
        var lowerLine = line.ToLowerInvariant();
        
        // Common header/footer patterns
        var skipPatterns = new[]
        {
            "page ", "date:", "revision:", "point list", "header", "footer",
            "name", "address", "type", "description", "===", "---", "___"
        };

        return skipPatterns.Any(pattern => lowerLine.Contains(pattern));
    }

    /// <summary>
    /// Parses a single line into a PointRecord.
    /// This is a placeholder implementation that should be customized based on actual data format.
    /// </summary>
    private PointRecord? ParseLine(string line)
    {
        // This is a simple implementation that splits by common delimiters
        // In a real implementation, this would need to match the actual format of the point list
        
        var separators = new[] { '\t', '|', ',' };
        var parts = line.Split(separators, StringSplitOptions.TrimEntries);

        // Need at least a name to create a record
        if (parts.Length < 1 || string.IsNullOrWhiteSpace(parts[0]))
        {
            return null;
        }

        var record = new PointRecord();

        // Map parts to record properties based on position
        // This mapping should be adjusted based on actual data format
        for (int i = 0; i < parts.Length; i++)
        {
            var part = parts[i].Trim();
            if (string.IsNullOrWhiteSpace(part))
                continue;

            switch (i)
            {
                case 0:
                    // First non-empty field is typically an address or index
                    if (int.TryParse(part, out _))
                        record.Address = part;
                    else
                        record.Name = part;
                    break;
                case 1:
                    if (string.IsNullOrEmpty(record.Name))
                        record.Name = part;
                    else if (string.IsNullOrEmpty(record.Address))
                        record.Address = part;
                    else
                        record.Type = part;
                    break;
                case 2:
                    if (string.IsNullOrEmpty(record.Name))
                        record.Name = part;
                    else if (string.IsNullOrEmpty(record.Type))
                        record.Type = part;
                    else
                        record.NormalState = part;
                    break;
                case 3:
                    if (string.IsNullOrEmpty(record.Type))
                        record.Type = part;
                    else
                        record.State1 = part;
                    break;
                case 4:
                    record.State0 = part;
                    break;
                case 5:
                    record.Alarm = part;
                    break;
                default:
                    // Store additional fields in AdditionalProperties
                    record.AdditionalProperties[$"Field{i}"] = part;
                    break;
            }
        }

        // Only return records that have at least a name
        return !string.IsNullOrWhiteSpace(record.Name) ? record : null;
    }
}
