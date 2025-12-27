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
        builder.Services.AddSingleton<BackupService>();
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

            // Check Office availability on startup
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

            await app.RunAsync();
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }
}
