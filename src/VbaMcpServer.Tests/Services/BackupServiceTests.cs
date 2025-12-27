using FluentAssertions;
using Microsoft.Extensions.Logging;
using Moq;
using VbaMcpServer.Services;

namespace VbaMcpServer.Tests.Services;

public class BackupServiceTests : IDisposable
{
    private readonly BackupService _service;
    private readonly string _testBackupDirectory;

    public BackupServiceTests()
    {
        var logger = new Mock<ILogger<BackupService>>();
        _service = new BackupService(logger.Object);

        _testBackupDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "backups"
        );
    }

    public void Dispose()
    {
        // Clean up test backups
        if (Directory.Exists(_testBackupDirectory))
        {
            foreach (var file in Directory.GetFiles(_testBackupDirectory, "TestWorkbook_*.bas"))
            {
                try { File.Delete(file); } catch { }
            }
        }
    }

    [Fact]
    public void BackupModule_CreatesBackupFile()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var code = "Sub Test()\n    MsgBox \"Hello\"\nEnd Sub";

        // Act
        var backupPath = _service.BackupModule(filePath, moduleName, code);

        // Assert
        File.Exists(backupPath).Should().BeTrue();
        var savedCode = File.ReadAllText(backupPath);
        savedCode.Should().Be(code);
        backupPath.Should().Contain("TestWorkbook_Module1_");
    }

    [Fact]
    public void BackupModule_ReturnsBackupPath()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var code = "Sub Test()\nEnd Sub";

        // Act
        var backupPath = _service.BackupModule(filePath, moduleName, code);

        // Assert
        backupPath.Should().NotBeNullOrEmpty();
        backupPath.Should().EndWith(".bas");
    }

    [Fact]
    public void ListBackups_ReturnsEmptyList_WhenNoBackups()
    {
        // Act
        var backups = _service.ListBackups("NonExistentFile.xlsm");

        // Assert
        backups.Should().BeEmpty();
    }

    [Fact]
    public void ListBackups_ReturnsBackups_WhenBackupsExist()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var code = "Sub Test()\nEnd Sub";
        _service.BackupModule(filePath, moduleName, code);

        // Act
        var backups = _service.ListBackups(filePath);

        // Assert
        backups.Should().NotBeEmpty();
        backups[0].FileName.Should().Contain("TestWorkbook_Module1_");
    }

    [Fact]
    public void RestoreBackup_ReturnsCode_WhenBackupExists()
    {
        // Arrange
        var filePath = "C:\\Test\\TestWorkbook.xlsm";
        var moduleName = "Module1";
        var code = "Sub Test()\n    MsgBox \"Restored\"\nEnd Sub";
        var backupPath = _service.BackupModule(filePath, moduleName, code);

        // Act
        var restoredCode = _service.RestoreBackup(backupPath);

        // Assert
        restoredCode.Should().Be(code);
    }

    [Fact]
    public void RestoreBackup_ThrowsException_WhenBackupDoesNotExist()
    {
        // Arrange
        var nonExistentPath = Path.Combine(_testBackupDirectory, "NonExistent.bas");

        // Act
        var act = () => _service.RestoreBackup(nonExistentPath);

        // Assert
        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void CleanupOldBackups_ReturnsZero_WhenNoOldBackups()
    {
        // Act
        var deleted = _service.CleanupOldBackups();

        // Assert
        deleted.Should().BeGreaterThanOrEqualTo(0);
    }
}
