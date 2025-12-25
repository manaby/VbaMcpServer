namespace VbaMcpServer.Models;

/// <summary>
/// Represents information about a VBA procedure
/// </summary>
public class ProcedureInfo
{
    /// <summary>
    /// Name of the procedure
    /// </summary>
    public required string Name { get; set; }

    /// <summary>
    /// Type of the procedure (Sub, Function, Property Get/Let/Set)
    /// </summary>
    public required string Type { get; set; }

    /// <summary>
    /// Starting line number (1-based)
    /// </summary>
    public int StartLine { get; set; }

    /// <summary>
    /// Number of lines in the procedure
    /// </summary>
    public int LineCount { get; set; }

    /// <summary>
    /// Access modifier (Public, Private, Friend)
    /// </summary>
    public string? AccessModifier { get; set; }
}

/// <summary>
/// VBA procedure types
/// </summary>
public enum VbaProcedureType
{
    Sub = 0,
    Function = 1,
    PropertyGet = 2,
    PropertyLet = 3,
    PropertySet = 4
}
