using Microsoft.Extensions.Logging;
using VbaMcpServer.Exceptions;

namespace VbaMcpServer.Services;

/// <summary>
/// Service for managing VBA code backups
/// </summary>
public class BackupService
{
    private readonly ILogger<BackupService> _logger;
    private readonly string _backupDirectory;
    private readonly int _retentionDays;

    public BackupService(ILogger<BackupService> logger)
    {
        _logger = logger;
        _backupDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "backups"
        );
        _retentionDays = 30;

        EnsureBackupDirectoryExists();
    }

    /// <summary>
    /// Create a backup of the module code before modification
    /// </summary>
    public string BackupModule(string filePath, string moduleName, string code)
    {
        var fileName = SanitizeFileName(Path.GetFileNameWithoutExtension(filePath));
        var safeName = SanitizeFileName(moduleName);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var backupFileName = $"{fileName}_{safeName}_{timestamp}.bas";
        var backupPath = Path.Combine(_backupDirectory, backupFileName);

        try
        {
            File.WriteAllText(backupPath, code);
            _logger.LogInformation("Backup created: {Path}", backupPath);
            return backupPath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to create backup for {Module}", moduleName);
            throw new BackupFailedException($"Failed to create backup for module '{moduleName}'", ex, filePath, moduleName);
        }
    }

    /// <summary>
    /// List all backups for a specific file
    /// </summary>
    public List<BackupInfo> ListBackups(string? filePath = null)
    {
        var result = new List<BackupInfo>();

        if (!Directory.Exists(_backupDirectory))
        {
            return result;
        }

        var searchPattern = filePath != null
            ? $"{SanitizeFileName(Path.GetFileNameWithoutExtension(filePath))}_*.bas"
            : "*.bas";

        foreach (var file in Directory.GetFiles(_backupDirectory, searchPattern))
        {
            var fileInfo = new FileInfo(file);
            result.Add(new BackupInfo
            {
                FileName = fileInfo.Name,
                FullPath = fileInfo.FullName,
                CreatedAt = fileInfo.CreationTime,
                SizeBytes = fileInfo.Length
            });
        }

        return result.OrderByDescending(b => b.CreatedAt).ToList();
    }

    /// <summary>
    /// Restore a backup
    /// </summary>
    public string RestoreBackup(string backupPath)
    {
        if (!File.Exists(backupPath))
        {
            throw new FileNotFoundException($"Backup file not found: {backupPath}");
        }

        return File.ReadAllText(backupPath);
    }

    /// <summary>
    /// Clean up old backups
    /// </summary>
    public int CleanupOldBackups()
    {
        if (!Directory.Exists(_backupDirectory))
        {
            return 0;
        }

        var cutoffDate = DateTime.Now.AddDays(-_retentionDays);
        var deleted = 0;

        foreach (var file in Directory.GetFiles(_backupDirectory, "*.bas"))
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.CreationTime < cutoffDate)
            {
                try
                {
                    File.Delete(file);
                    deleted++;
                    _logger.LogDebug("Deleted old backup: {Path}", file);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to delete old backup: {Path}", file);
                }
            }
        }

        if (deleted > 0)
        {
            _logger.LogInformation("Cleaned up {Count} old backups", deleted);
        }

        return deleted;
    }

    private void EnsureBackupDirectoryExists()
    {
        if (!Directory.Exists(_backupDirectory))
        {
            Directory.CreateDirectory(_backupDirectory);
            _logger.LogInformation("Created backup directory: {Path}", _backupDirectory);
        }
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Join("_", name.Split(invalid, StringSplitOptions.RemoveEmptyEntries));
    }
}

/// <summary>
/// Information about a backup file
/// </summary>
public class BackupInfo
{
    public required string FileName { get; set; }
    public required string FullPath { get; set; }
    public DateTime CreatedAt { get; set; }
    public long SizeBytes { get; set; }
}
