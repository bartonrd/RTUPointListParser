namespace RTUPointlistParse.Models;

/// <summary>
/// Represents a single point record from a point list.
/// </summary>
public class PointRecord
{
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string Unit { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string NormalState { get; set; } = string.Empty;
    public string State1 { get; set; } = string.Empty;
    public string State0 { get; set; } = string.Empty;
    public string Alarm { get; set; } = string.Empty;
    public string EmsNumber { get; set; } = string.Empty;
    public Dictionary<string, string> AdditionalProperties { get; set; } = new();

    public override string ToString()
    {
        return $"PointRecord: {Name} [{Address}] - {Type}";
    }
}
