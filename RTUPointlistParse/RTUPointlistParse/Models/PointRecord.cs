namespace RTUPointlistParse.Models;

/// <summary>
/// Represents a single point record from a point list.
/// </summary>
public class PointRecord
{
    /// <summary>
    /// The name or identifier of the point.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// The address of the point (e.g., register address, DNP address).
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// The type of the point (e.g., AI, AO, DI, DO, Counter).
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>
    /// The unit of measurement for the point (e.g., kW, V, A, °F).
    /// </summary>
    public string Unit { get; set; } = string.Empty;

    /// <summary>
    /// A description of the point.
    /// </summary>
    public string Description { get; set; } = string.Empty;
}
