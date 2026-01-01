namespace VbaMcpServer.Models;

/// <summary>
/// Information about a control in an Access form
/// </summary>
public class FormControlInfo
{
    public required string Name { get; init; }
    public required string ControlType { get; init; }
    public required int ControlTypeId { get; init; }
    public required string Section { get; init; }
    public required int SectionId { get; init; }
    public required int Left { get; init; }
    public required int Top { get; init; }
    public required int Width { get; init; }
    public required int Height { get; init; }
    public required bool Visible { get; init; }
    public bool? Enabled { get; init; }
    public int? TabIndex { get; init; }
    public string? ControlSource { get; init; }
    public string? Parent { get; init; }
    public string? SourceObject { get; init; }
}
