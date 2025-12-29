namespace VbaMcpServer.Models;

/// <summary>
/// Result of a table data query
/// </summary>
public class TableDataResult
{
    /// <summary>
    /// List of column names
    /// </summary>
    public required List<string> ColumnNames { get; init; }

    /// <summary>
    /// List of rows, each row is a dictionary of column name to value
    /// </summary>
    public required List<Dictionary<string, object?>> Rows { get; init; }

    /// <summary>
    /// Total number of rows in the table (before pagination)
    /// </summary>
    public int TotalRows { get; init; }

    /// <summary>
    /// Number of rows returned in this result
    /// </summary>
    public int ReturnedRows { get; init; }

    /// <summary>
    /// Whether there are more rows available
    /// </summary>
    public bool HasMore { get; init; }
}
