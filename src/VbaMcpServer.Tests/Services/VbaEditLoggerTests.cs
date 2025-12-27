using System.Text.Json;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using Moq;
using VbaMcpServer.Logging;

namespace VbaMcpServer.Tests.Services;

public class VbaEditLoggerTests : IDisposable
{
    private readonly VbaEditLogger _logger;
    private readonly string _testLogDirectory;

    public VbaEditLoggerTests()
    {
        var mockLogger = new Mock<ILogger<VbaEditLogger>>();
        _logger = new VbaEditLogger(mockLogger.Object);

        _testLogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "logs",
            "vba"
        );
    }

    public void Dispose()
    {
        // Clean up test log files
        if (Directory.Exists(_testLogDirectory))
        {
            var todayLogFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
            if (File.Exists(todayLogFile))
            {
                try { File.Delete(todayLogFile); } catch { }
            }
        }
    }

    [Fact]
    public void LogModuleRead_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var lineCount = 10;

        // Act
        _logger.LogModuleRead(filePath, moduleName, lineCount);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        File.Exists(logFile).Should().BeTrue();

        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("ModuleRead");
        logContent.Should().Contain(moduleName);
    }

    [Fact]
    public void LogModuleWritten_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var lineCount = 15;
        var backupPath = "C:\\Backup\\backup.bas";

        // Act
        _logger.LogModuleWritten(filePath, moduleName, lineCount, backupPath);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("ModuleWritten");
        logContent.Should().Contain(backupPath.Replace("\\", "\\\\"));  // JSON escapes backslashes
    }

    [Fact]
    public void LogModuleCreated_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "NewModule";
        var moduleType = "StdModule";

        // Act
        _logger.LogModuleCreated(filePath, moduleName, moduleType);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("ModuleCreated");
        logContent.Should().Contain(moduleType);
    }

    [Fact]
    public void LogModuleDeleted_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "OldModule";
        var backupPath = "C:\\Backup\\backup.bas";

        // Act
        _logger.LogModuleDeleted(filePath, moduleName, backupPath);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("ModuleDeleted");
        logContent.Should().Contain(backupPath.Replace("\\", "\\\\"));  // JSON escapes backslashes
    }

    [Fact]
    public void LogBackupCreated_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var backupPath = "C:\\Backup\\backup.bas";

        // Act
        _logger.LogBackupCreated(filePath, moduleName, backupPath);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("BackupCreated");
        logContent.Should().Contain(backupPath.Replace("\\", "\\\\"));  // JSON escapes backslashes
    }

    [Fact]
    public void LogModuleExported_CreatesLogEntry()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var outputPath = "C:\\Export\\module.bas";

        // Act
        _logger.LogModuleExported(filePath, moduleName, outputPath);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var logContent = File.ReadAllText(logFile);
        logContent.Should().Contain("ModuleExported");
        logContent.Should().Contain(outputPath.Replace("\\", "\\\\"));  // JSON escapes backslashes
    }

    [Fact]
    public void LogEntries_AreValidJson()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";

        // Act
        _logger.LogModuleRead(filePath, moduleName, 10);

        // Assert
        var logFile = Path.Combine(_testLogDirectory, $"vba-{DateTime.Now:yyyy-MM-dd}.log");
        var lines = File.ReadAllLines(logFile);

        foreach (var line in lines)
        {
            if (!string.IsNullOrWhiteSpace(line))
            {
                var act = () => JsonSerializer.Deserialize<LogEntry>(line);
                act.Should().NotThrow();
            }
        }
    }
}
