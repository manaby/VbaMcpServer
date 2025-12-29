namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a query is not found in the database
/// </summary>
public class QueryNotFoundException : VbaMcpException
{
    public QueryNotFoundException(string queryName, string? filePath = null)
        : base($"Query not found: {queryName}", "QUERY_NOT_FOUND", filePath)
    {
    }
}
