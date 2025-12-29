namespace VbaMcpServer.Models;

/// <summary>
/// Result of an action query execution
/// </summary>
public class QueryExecutionResult
{
    /// <summary>
    /// Whether the query executed successfully
    /// </summary>
    public bool Success { get; init; }

    /// <summary>
    /// Number of records affected by the query
    /// </summary>
    public int RecordsAffected { get; init; }

    /// <summary>
    /// Optional message (e.g., error message)
    /// </summary>
    public string? Message { get; init; }
}
