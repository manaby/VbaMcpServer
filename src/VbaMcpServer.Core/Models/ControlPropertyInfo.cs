namespace VbaMcpServer.Models;

/// <summary>
/// Detailed property information for a control
/// </summary>
public class ControlPropertyInfo
{
    public required string File { get; init; }
    public required string ObjectName { get; init; }
    public required string ControlName { get; init; }
    public required string ControlType { get; init; }
    public required Dictionary<string, object?> Properties { get; init; }
}
