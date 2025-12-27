using System.Text.Json;
using VbaMcpServer.Logging;

namespace VbaMcpServer.GUI.Services;

public class LogViewerService : IDisposable
{
    private FileSystemWatcher? _watcher;
    private readonly string _logDirectory;

    public event EventHandler<string>? NewLogLine;

    public LogViewerService()
    {
        _logDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "logs"
        );
    }

    public List<string> LoadServerLogs(DateTime? date = null)
    {
        var targetDate = date ?? DateTime.Now;
        var logFile = Path.Combine(_logDirectory, "server", $"server-{targetDate:yyyy-MM-dd}.log");
        return LoadLogFile(logFile);
    }

    public List<string> LoadVbaEditLogs(DateTime? date = null)
    {
        var targetDate = date ?? DateTime.Now;
        var logFile = Path.Combine(_logDirectory, "vba", $"vba-{targetDate:yyyy-MM-dd}.log");
        return LoadLogFile(logFile);
    }

    private List<string> LoadLogFile(string filePath)
    {
        var logs = new List<string>();

        if (!File.Exists(filePath))
        {
            return logs;
        }

        try
        {
            using var reader = new StreamReader(new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                logs.Add(FormatLogLine(line));
            }
        }
        catch (Exception ex)
        {
            logs.Add($"Error reading log file: {ex.Message}");
        }

        return logs;
    }

    private string FormatLogLine(string jsonLine)
    {
        try
        {
            using var doc = JsonDocument.Parse(jsonLine);
            var root = doc.RootElement;

            var timestamp = root.TryGetProperty("@t", out var t) ? t.GetString() :
                           root.TryGetProperty("Timestamp", out var ts) ? ts.GetString() :
                           DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            var level = root.TryGetProperty("@l", out var l) ? l.GetString() :
                       root.TryGetProperty("Level", out var lv) ? lv.GetString() :
                       "INFO";

            var message = root.TryGetProperty("@m", out var m) ? m.GetString() :
                         root.TryGetProperty("@mt", out var mt) ? mt.GetString() :
                         root.TryGetProperty("EventType", out var et) ? $"{et.GetString()} - {GetEventDetails(root)}" :
                         jsonLine;

            return $"{timestamp} [{level}] {message}";
        }
        catch
        {
            return jsonLine;
        }
    }

    private string GetEventDetails(JsonElement root)
    {
        var details = new List<string>();

        if (root.TryGetProperty("FilePath", out var fp))
            details.Add($"File: {Path.GetFileName(fp.GetString())}");

        if (root.TryGetProperty("ModuleName", out var mn))
            details.Add($"Module: {mn.GetString()}");

        if (root.TryGetProperty("LineCount", out var lc))
            details.Add($"Lines: {lc.GetInt32()}");

        return string.Join(", ", details);
    }

    public void StartWatching(string logType, Action<string> onNewLine)
    {
        var watchPath = logType.ToLower() == "server"
            ? Path.Combine(_logDirectory, "server")
            : Path.Combine(_logDirectory, "vba");

        if (!Directory.Exists(watchPath))
        {
            Directory.CreateDirectory(watchPath);
        }

        _watcher?.Dispose();
        _watcher = new FileSystemWatcher(watchPath)
        {
            Filter = "*.log",
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
        };

        _watcher.Changed += (sender, e) =>
        {
            try
            {
                Thread.Sleep(100);
                using var reader = new StreamReader(new FileStream(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
                reader.BaseStream.Seek(0, SeekOrigin.End);
                reader.BaseStream.Seek(Math.Max(0, reader.BaseStream.Length - 4096), SeekOrigin.Begin);

                string? line;
                var lines = new List<string>();
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                }

                if (lines.Count > 0)
                {
                    var lastLine = FormatLogLine(lines[^1]);
                    NewLogLine?.Invoke(this, lastLine);
                    onNewLine?.Invoke(lastLine);
                }
            }
            catch
            {
                // Ignore file access errors
            }
        };

        _watcher.EnableRaisingEvents = true;
    }

    public void StopWatching()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Dispose();
            _watcher = null;
        }
    }

    public void SaveLogsToFile(string outputPath, List<string> logs)
    {
        File.WriteAllLines(outputPath, logs);
    }

    public void Dispose()
    {
        StopWatching();
    }
}
