using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using VbaMcpServer.Services;

namespace VbaMcpServer;

class Program
{
    static async Task Main(string[] args)
    {
        var builder = Host.CreateEmptyApplicationBuilder(settings: null);

        // Configure logging
        builder.Logging.AddConsole(options =>
        {
            options.LogToStandardErrorThreshold = LogLevel.Trace;
        });
        builder.Logging.SetMinimumLevel(LogLevel.Information);

        // Register services
        builder.Services.AddSingleton<ExcelComService>();
        builder.Services.AddSingleton<AccessComService>();
        builder.Services.AddSingleton<BackupService>();

        // Configure MCP Server
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        var app = builder.Build();

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
}
