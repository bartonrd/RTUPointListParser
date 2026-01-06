using RTUPointlistParse.Models;
using System.Text.RegularExpressions;

namespace RTUPointlistParse.Services;

/// <summary>
/// Service for parsing point list data from text content.
/// </summary>
public class PointListParser
{
    /// <summary>
    /// Parses text content and extracts point records.
    /// </summary>
    /// <param name="content">Text content to parse.</param>
    /// <returns>List of parsed point records.</returns>
    public List<PointRecord> ParsePointList(string content)
    {
        var records = new List<PointRecord>();
        var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        // Detect if this is a CSV format (check if first data line has commas)
        bool isCSV = DetectCSVFormat(lines);

        foreach (var line in lines)
        {
            // Skip empty lines or lines with only whitespace
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Skip common header/footer patterns
            if (IsHeaderOrFooter(line))
                continue;

            try
            {
                var record = isCSV ? ParseCSVLine(line) : ParseLine(line);
                if (record != null)
                {
                    records.Add(record);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to parse line: {line}");
                Console.WriteLine($"  Error: {ex.Message}");
            }
        }

        return records;
    }

    /// <summary>
    /// Detects if content is in CSV format.
    /// </summary>
    private bool DetectCSVFormat(string[] lines)
    {
        // Look at first few non-empty, non-header lines
        foreach (var line in lines.Take(10))
        {
            if (string.IsNullOrWhiteSpace(line) || IsHeaderOrFooter(line))
                continue;

            // If line has commas, likely CSV
            if (line.Contains(','))
                return true;

            break;
        }

        return false;
    }

    /// <summary>
    /// Checks if a line is likely a header or footer.
    /// </summary>
    private bool IsHeaderOrFooter(string line)
    {
        var lowerLine = line.ToLower().Trim();
        
        // Common header/footer patterns
        var patterns = new[]
        {
            "page", "date:", "revision", "project", "document",
            "===", "---", "___", "point list"
        };

        // Check if line starts with common header column names and contains others
        // This avoids false positives for data rows
        var startsWithName = lowerLine.StartsWith("name");
        var containsAddress = lowerLine.Contains("address");
        var containsType = lowerLine.Contains("type");
        
        if (startsWithName && containsAddress && containsType)
        {
            return true;
        }

        return patterns.Any(pattern => lowerLine.Contains(pattern));
    }

    /// <summary>
    /// Parses a CSV line into a PointRecord.
    /// </summary>
    private PointRecord? ParseCSVLine(string line)
    {
        // Simple CSV parser - handles quoted fields
        var parts = SplitCSV(line);

        // Need at least a name to create a record
        if (parts.Count == 0 || string.IsNullOrWhiteSpace(parts[0]))
            return null;

        var record = new PointRecord
        {
            Name = parts.Count > 0 ? parts[0].Trim() : string.Empty,
            Address = parts.Count > 1 ? parts[1].Trim() : string.Empty,
            Type = parts.Count > 2 ? parts[2].Trim() : string.Empty,
            Unit = parts.Count > 3 ? parts[3].Trim() : string.Empty,
            Description = parts.Count > 4 ? parts[4].Trim() : string.Empty
        };

        return record;
    }

    /// <summary>
    /// Splits a CSV line respecting quoted fields.
    /// </summary>
    private List<string> SplitCSV(string line)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                // Check if this is an escaped quote ("")
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    // Add a single quote to the output and skip the next quote
                    current.Append('"');
                    i++;
                }
                else
                {
                    // Toggle quote state
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        result.Add(current.ToString());
        return result;
    }

    /// <summary>
    /// Parses a single line into a PointRecord.
    /// </summary>
    private PointRecord? ParseLine(string line)
    {
        // Split by tabs first, then by multiple spaces (2 or more)
        var parts = Regex.Split(line.Trim(), @"\t+|\s{2,}")
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToArray();

        // Need at least a name to create a record
        if (parts.Length == 0)
            return null;

        // Basic parsing - customize based on actual point list format
        var record = new PointRecord
        {
            Name = parts.Length > 0 ? parts[0].Trim() : string.Empty,
            Address = parts.Length > 1 ? parts[1].Trim() : string.Empty,
            Type = parts.Length > 2 ? parts[2].Trim() : string.Empty,
            Unit = parts.Length > 3 ? parts[3].Trim() : string.Empty,
            Description = parts.Length > 4 ? string.Join(" ", parts.Skip(4).Select(p => p.Trim())) : string.Empty
        };

        return record;
    }
}
