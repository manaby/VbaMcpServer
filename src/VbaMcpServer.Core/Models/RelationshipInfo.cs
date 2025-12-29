namespace VbaMcpServer.Models;

/// <summary>
/// Represents a relationship between two tables in an Access database
/// </summary>
public class RelationshipInfo
{
    /// <summary>
    /// Name of the relationship
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Name of the parent (primary) table
    /// </summary>
    public required string ParentTable { get; init; }

    /// <summary>
    /// Name of the child (foreign) table
    /// </summary>
    public required string ChildTable { get; init; }

    /// <summary>
    /// Parent table field(s) in the relationship
    /// </summary>
    public required List<string> ParentFields { get; init; }

    /// <summary>
    /// Child table field(s) in the relationship
    /// </summary>
    public required List<string> ChildFields { get; init; }

    /// <summary>
    /// Whether referential integrity is enforced
    /// </summary>
    public bool EnforceReferentialIntegrity { get; init; }

    /// <summary>
    /// Whether cascade updates are enabled
    /// </summary>
    public bool CascadeUpdates { get; init; }

    /// <summary>
    /// Whether cascade deletes are enabled
    /// </summary>
    public bool CascadeDeletes { get; init; }

    /// <summary>
    /// Relationship type (One-to-One, One-to-Many, etc.)
    /// </summary>
    public required string RelationType { get; init; }
}
