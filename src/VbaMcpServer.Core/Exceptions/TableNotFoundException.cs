namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a table is not found in the database
/// </summary>
public class TableNotFoundException : VbaMcpException
{
    public TableNotFoundException(string tableName, string? filePath = null)
        : base($"Table not found: {tableName}", "TABLE_NOT_FOUND", filePath)
    {
    }
}
