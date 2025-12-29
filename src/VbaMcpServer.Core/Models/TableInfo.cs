namespace VbaMcpServer.Models;

/// <summary>
/// Information about an Access table
/// </summary>
public class TableInfo
{
    /// <summary>
    /// Table name
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Table type (e.g., "Table", "LinkedTable", "System")
    /// </summary>
    public required string Type { get; init; }

    /// <summary>
    /// Number of records in the table
    /// </summary>
    public int RecordCount { get; init; }

    /// <summary>
    /// Date the table was created
    /// </summary>
    public DateTime? DateCreated { get; init; }

    /// <summary>
    /// Date the table was last modified
    /// </summary>
    public DateTime? DateModified { get; init; }
}
