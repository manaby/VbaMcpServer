using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace VbaMcpServer.Logging;

/// <summary>
/// Logger for VBA edit operations, writing to a separate JSON log file
/// </summary>
public class VbaEditLogger : IVbaEditLogger
{
    private readonly ILogger<VbaEditLogger> _logger;
    private readonly string _logDirectory;
    private readonly object _fileLock = new();

    public VbaEditLogger(ILogger<VbaEditLogger> logger)
    {
        _logger = logger;
        _logDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "logs",
            "vba"
        );

        EnsureLogDirectoryExists();
    }

    public void LogModuleRead(string filePath, string moduleName, int lineCount)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "ModuleRead",
            FilePath = filePath,
            ModuleName = moduleName,
            LineCount = lineCount
        });
    }

    public void LogModuleWritten(string filePath, string moduleName, int lineCount, string? backupPath = null)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "ModuleWritten",
            FilePath = filePath,
            ModuleName = moduleName,
            LineCount = lineCount,
            BackupPath = backupPath
        });
    }

    public void LogModuleCreated(string filePath, string moduleName, string moduleType)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "ModuleCreated",
            FilePath = filePath,
            ModuleName = moduleName,
            ModuleType = moduleType
        });
    }

    public void LogModuleDeleted(string filePath, string moduleName, string backupPath)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "ModuleDeleted",
            FilePath = filePath,
            ModuleName = moduleName,
            BackupPath = backupPath
        });
    }

    public void LogBackupCreated(string filePath, string moduleName, string backupPath)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "BackupCreated",
            FilePath = filePath,
            ModuleName = moduleName,
            BackupPath = backupPath
        });
    }

    public void LogModuleExported(string filePath, string moduleName, string outputPath)
    {
        WriteLogEntry(new LogEntry
        {
            Timestamp = DateTime.Now,
            EventType = "ModuleExported",
            FilePath = filePath,
            ModuleName = moduleName,
            AdditionalInfo = outputPath
        });
    }

    private void WriteLogEntry(LogEntry entry)
    {
        try
        {
            var logFilePath = GetLogFilePath();
            var jsonLine = JsonSerializer.Serialize(entry, new JsonSerializerOptions
            {
                WriteIndented = false
            });

            lock (_fileLock)
            {
                File.AppendAllText(logFilePath, jsonLine + Environment.NewLine);
            }

            _logger.LogDebug("VBA edit logged: {EventType} - {Module}", entry.EventType, entry.ModuleName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to write VBA edit log entry");
        }
    }

    private string GetLogFilePath()
    {
        var date = DateTime.Now.ToString("yyyy-MM-dd");
        return Path.Combine(_logDirectory, $"vba-{date}.log");
    }

    private void EnsureLogDirectoryExists()
    {
        if (!Directory.Exists(_logDirectory))
        {
            Directory.CreateDirectory(_logDirectory);
            _logger.LogInformation("Created VBA log directory: {Path}", _logDirectory);
        }
    }
}
