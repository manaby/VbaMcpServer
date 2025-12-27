namespace VbaMcpServer.Logging;

/// <summary>
/// Interface for logging VBA edit operations
/// </summary>
public interface IVbaEditLogger
{
    /// <summary>
    /// Log when a module is read
    /// </summary>
    void LogModuleRead(string filePath, string moduleName, int lineCount);

    /// <summary>
    /// Log when a module is written
    /// </summary>
    void LogModuleWritten(string filePath, string moduleName, int lineCount, string? backupPath = null);

    /// <summary>
    /// Log when a module is created
    /// </summary>
    void LogModuleCreated(string filePath, string moduleName, string moduleType);

    /// <summary>
    /// Log when a module is deleted
    /// </summary>
    void LogModuleDeleted(string filePath, string moduleName, string backupPath);

    /// <summary>
    /// Log when a backup is created
    /// </summary>
    void LogBackupCreated(string filePath, string moduleName, string backupPath);

    /// <summary>
    /// Log when a module is exported
    /// </summary>
    void LogModuleExported(string filePath, string moduleName, string outputPath);
}
