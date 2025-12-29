namespace VbaMcpServer.Models;

/// <summary>
/// Information about a saved Access query
/// </summary>
public class QueryInfo
{
    /// <summary>
    /// Query name
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Query type (e.g., "Select", "Action", "Crosstab", "Delete", "Update", "Append")
    /// </summary>
    public required string QueryType { get; init; }

    /// <summary>
    /// Date the query was created
    /// </summary>
    public DateTime? DateCreated { get; init; }

    /// <summary>
    /// Date the query was last modified
    /// </summary>
    public DateTime? DateModified { get; init; }

    /// <summary>
    /// Number of parameters in the query
    /// </summary>
    public int ParameterCount { get; init; }
}
