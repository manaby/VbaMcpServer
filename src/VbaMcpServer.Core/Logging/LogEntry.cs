namespace VbaMcpServer.Logging;

/// <summary>
/// Represents a VBA edit log entry
/// </summary>
public class LogEntry
{
    public DateTime Timestamp { get; set; }
    public required string EventType { get; set; }
    public required string FilePath { get; set; }
    public required string ModuleName { get; set; }
    public int? LineCount { get; set; }
    public string? BackupPath { get; set; }
    public string? ModuleType { get; set; }
    public string? AdditionalInfo { get; set; }
}
