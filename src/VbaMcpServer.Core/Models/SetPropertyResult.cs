namespace VbaMcpServer.Models;

/// <summary>
/// Result of a property set operation
/// </summary>
public class SetPropertyResult
{
    public required bool Success { get; init; }
    public string? File { get; init; }
    public string? ObjectName { get; init; }
    public string? ControlName { get; init; }
    public string? PropertyName { get; init; }
    public object? PreviousValue { get; init; }
    public object? NewValue { get; init; }
    public string? Error { get; init; }
    public string? ErrorCode { get; init; }
}
