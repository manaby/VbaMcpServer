namespace VbaMcpServer.Models;

/// <summary>
/// Represents an index on an Access table
/// </summary>
public class IndexInfo
{
    /// <summary>
    /// Name of the index
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Fields that make up the index
    /// </summary>
    public required List<string> Fields { get; init; }

    /// <summary>
    /// Whether this is a primary key index
    /// </summary>
    public bool IsPrimary { get; init; }

    /// <summary>
    /// Whether this index enforces uniqueness
    /// </summary>
    public bool IsUnique { get; init; }

    /// <summary>
    /// Whether this is a foreign key index
    /// </summary>
    public bool IsForeign { get; init; }

    /// <summary>
    /// Whether NULL values are ignored in the index
    /// </summary>
    public bool IgnoreNulls { get; init; }

    /// <summary>
    /// Whether this index is required (cannot be deleted)
    /// </summary>
    public bool IsRequired { get; init; }
}
