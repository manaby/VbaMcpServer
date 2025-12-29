namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when attempting to create a query that already exists
/// </summary>
public class QueryAlreadyExistsException : VbaMcpException
{
    public QueryAlreadyExistsException(string queryName, string? filePath = null)
        : base($"Query already exists: {queryName}. Use replaceIfExists=true to overwrite.",
               "QUERY_ALREADY_EXISTS", filePath)
    {
    }
}
