using RTUPointlistParse.Models;

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
                var record = ParseLine(line);
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
    /// Checks if a line is likely a header or footer.
    /// </summary>
    private bool IsHeaderOrFooter(string line)
    {
        var lowerLine = line.ToLower().Trim();
        
        // Common header/footer patterns
        var patterns = new[]
        {
            "page", "date:", "revision", "project", "document",
            "===", "---", "___", "point list", "point name",
            "address", "type", "unit", "description"
        };

        return patterns.Any(pattern => lowerLine.Contains(pattern));
    }

    /// <summary>
    /// Parses a single line into a PointRecord.
    /// </summary>
    private PointRecord? ParseLine(string line)
    {
        // Normalize whitespace (tabs to spaces, multiple spaces to single)
        var normalizedLine = System.Text.RegularExpressions.Regex.Replace(line.Trim(), @"\s+", " ");

        // Split by multiple spaces or tabs to handle column-based formats
        var parts = normalizedLine.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

        // Need at least a name to create a record
        if (parts.Length == 0)
            return null;

        // Basic parsing - customize based on actual point list format
        var record = new PointRecord
        {
            Name = parts.Length > 0 ? parts[0] : string.Empty,
            Address = parts.Length > 1 ? parts[1] : string.Empty,
            Type = parts.Length > 2 ? parts[2] : string.Empty,
            Unit = parts.Length > 3 ? parts[3] : string.Empty,
            Description = parts.Length > 4 ? string.Join(" ", parts.Skip(4)) : string.Empty
        };

        return record;
    }
}
