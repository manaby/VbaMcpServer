namespace VbaMcpServer.Models;

/// <summary>
/// Represents a database object (form, report, module, etc.) in an Access database
/// </summary>
public class DatabaseObjectInfo
{
    /// <summary>
    /// Name of the object
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Type of the object (Form, Report, Module, etc.)
    /// </summary>
    public required string Type { get; init; }

    /// <summary>
    /// Date the object was created
    /// </summary>
    public DateTime? DateCreated { get; init; }

    /// <summary>
    /// Date the object was last modified
    /// </summary>
    public DateTime? DateModified { get; init; }

    /// <summary>
    /// Whether the object is loaded/open
    /// </summary>
    public bool IsLoaded { get; init; }
}
