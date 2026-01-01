using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Serilog;
using VbaMcpServer.Logging;
using VbaMcpServer.Services;

namespace VbaMcpServer;

class Program
{
    static async Task Main(string[] args)
    {
        var builder = Host.CreateEmptyApplicationBuilder(settings: null);

        // Configure configuration sources
        builder.Configuration.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

        // Configure Serilog
        var serverLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".vba-mcp-server",
            "logs",
            "server",
            $"server-{DateTime.Now:yyyy-MM-dd}.log"
        );

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .MinimumLevel.Override("Microsoft", Serilog.Events.LogEventLevel.Warning)
            .MinimumLevel.Override("System", Serilog.Events.LogEventLevel.Warning)
            .WriteTo.Console(
                outputTemplate: "{Timestamp:HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                standardErrorFromLevel: Serilog.Events.LogEventLevel.Verbose)
            .WriteTo.File(
                serverLogPath,
                rollingInterval: RollingInterval.Day,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        builder.Logging.ClearProviders();
        builder.Logging.AddSerilog(Log.Logger);

        // Register services
        builder.Services.AddSingleton<ExcelComService>();
        builder.Services.AddSingleton<AccessComService>();
        builder.Services.AddSingleton<IVbaEditLogger, VbaEditLogger>();

        // Configure MCP Server
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        var app = builder.Build();

        try
        {
            // Log startup
            var logger = app.Services.GetRequiredService<ILogger<Program>>();
            logger.LogInformation("VBA MCP Server starting...");
            logger.LogInformation("Version: {Version}", typeof(Program).Assembly.GetName().Version);

            // Read target file from environment variable (set by GUI)
            var targetFilePath = Environment.GetEnvironmentVariable("VBA_TARGET_FILE");

            // Check Office availability based on target file type
            if (!string.IsNullOrEmpty(targetFilePath))
            {
                logger.LogInformation("Target file: {FilePath}", targetFilePath);

                // ターゲットファイルが指定されている場合、そのファイルタイプに応じてチェック
                if (IsExcelFile(targetFilePath))
                {
                    try
                    {
                        var excelService = app.Services.GetRequiredService<ExcelComService>();
                        if (!excelService.IsExcelAvailable())
                        {
                            logger.LogWarning("Excel is not available or not running");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogWarning(ex, "Could not check Excel availability");
                    }
                }
                else if (IsAccessFile(targetFilePath))
                {
                    try
                    {
                        var accessService = app.Services.GetRequiredService<AccessComService>();
                        if (!accessService.IsAccessAvailable())
                        {
                            logger.LogWarning("Access is not available or not running");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogWarning(ex, "Could not check Access availability");
                    }
                }
                else
                {
                    logger.LogWarning("Unknown file type: {FilePath}", targetFilePath);
                }
            }
            else
            {
                // ターゲットファイルが指定されていない場合は、両方チェック（既存動作）
                try
                {
                    var excelService = app.Services.GetRequiredService<ExcelComService>();
                    if (!excelService.IsExcelAvailable())
                    {
                        logger.LogInformation("Excel is not available or not running");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Could not check Excel availability");
                }

                try
                {
                    var accessService = app.Services.GetRequiredService<AccessComService>();
                    if (!accessService.IsAccessAvailable())
                    {
                        logger.LogInformation("Access is not available or not running");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Could not check Access availability");
                }
            }

            await app.RunAsync();
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    private static bool IsExcelFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension is ".xlsm" or ".xlsx" or ".xlsb" or ".xls" or ".xltm";
    }

    private static bool IsAccessFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension is ".accdb" or ".mdb";
    }
}
