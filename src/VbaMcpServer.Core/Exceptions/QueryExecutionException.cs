namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when query execution fails
/// </summary>
public class QueryExecutionException : VbaMcpException
{
    public QueryExecutionException(string message, Exception innerException, string? filePath = null)
        : base($"Query execution failed: {message}", "QUERY_EXECUTION_ERROR", innerException, filePath)
    {
    }

    public QueryExecutionException(string message, string? filePath = null)
        : base($"Query execution failed: {message}", "QUERY_EXECUTION_ERROR", filePath)
    {
    }
}
