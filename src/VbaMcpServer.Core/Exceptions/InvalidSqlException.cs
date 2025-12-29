namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when invalid SQL is detected
/// </summary>
public class InvalidSqlException : VbaMcpException
{
    public InvalidSqlException(string message, string? filePath = null)
        : base($"Invalid SQL: {message}", "INVALID_SQL", filePath)
    {
    }
}
