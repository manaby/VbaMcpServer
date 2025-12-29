namespace VbaMcpServer.Models;

/// <summary>
/// Represents summary information about an Access database
/// </summary>
public class DatabaseInfo
{
    /// <summary>
    /// Full path to the database file
    /// </summary>
    public required string FilePath { get; init; }

    /// <summary>
    /// Database file size in bytes
    /// </summary>
    public long FileSizeBytes { get; init; }

    /// <summary>
    /// Database file size formatted as human-readable string
    /// </summary>
    public required string FileSizeFormatted { get; init; }

    /// <summary>
    /// Access database version
    /// </summary>
    public required string Version { get; init; }

    /// <summary>
    /// Number of tables in the database
    /// </summary>
    public int TableCount { get; init; }

    /// <summary>
    /// Number of queries in the database
    /// </summary>
    public int QueryCount { get; init; }

    /// <summary>
    /// Number of forms in the database
    /// </summary>
    public int FormCount { get; init; }

    /// <summary>
    /// Number of reports in the database
    /// </summary>
    public int ReportCount { get; init; }

    /// <summary>
    /// Number of relationships in the database
    /// </summary>
    public int RelationshipCount { get; init; }

    /// <summary>
    /// Date the database file was created
    /// </summary>
    public DateTime? DateCreated { get; init; }

    /// <summary>
    /// Date the database file was last modified
    /// </summary>
    public DateTime? DateModified { get; init; }

    /// <summary>
    /// Whether the database is password-protected
    /// </summary>
    public bool IsPasswordProtected { get; init; }
}
