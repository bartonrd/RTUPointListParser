using System.Collections.Generic;

namespace PointListProcessor;

/// <summary>
/// Represents a point in the point list with its associated properties.
/// </summary>
public class Point
{
    /// <summary>
    /// Gets or sets the point name or identifier.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the point type (e.g., Analog Input, Digital Output, etc.).
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the point address or index.
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the point description.
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets any additional properties or metadata.
    /// </summary>
    public Dictionary<string, string> AdditionalProperties { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// Returns a string representation of the point.
    /// </summary>
    public override string ToString()
    {
        return $"{Name} | {Type} | {Address} | {Description}";
    }
}
