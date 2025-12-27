namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when backup creation fails
/// </summary>
public class BackupFailedException : VbaMcpException
{
    public BackupFailedException(string message, string? filePath = null, string? moduleName = null)
        : base(message, "BACKUP_FAILED", filePath, moduleName)
    {
    }

    public BackupFailedException(string message, Exception innerException,
                                string? filePath = null, string? moduleName = null)
        : base(message, "BACKUP_FAILED", innerException, filePath, moduleName)
    {
    }
}
