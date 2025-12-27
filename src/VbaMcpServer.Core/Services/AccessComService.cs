using Microsoft.Extensions.Logging;
using VbaMcpServer.Models;

namespace VbaMcpServer.Services;

/// <summary>
/// Service for interacting with Access VBA projects via COM
/// </summary>
public class AccessComService
{
    private readonly ILogger<AccessComService> _logger;

    public AccessComService(ILogger<AccessComService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Check if Access is available
    /// </summary>
    public bool IsAccessAvailable()
    {
        // TODO: Implement Access availability check
        _logger.LogDebug("Access availability check not yet implemented");
        return false;
    }

    /// <summary>
    /// List all open Access databases
    /// </summary>
    public List<string> ListOpenDatabases()
    {
        // TODO: Implement
        _logger.LogDebug("ListOpenDatabases not yet implemented");
        return new List<string>();
    }

    /// <summary>
    /// List all modules in a database
    /// </summary>
    public List<ModuleInfo> ListModules(string filePath)
    {
        // TODO: Implement
        _logger.LogDebug("ListModules for Access not yet implemented");
        throw new NotImplementedException("Access support coming soon");
    }

    /// <summary>
    /// Read the code from a module
    /// </summary>
    public string ReadModule(string filePath, string moduleName)
    {
        // TODO: Implement
        throw new NotImplementedException("Access support coming soon");
    }

    /// <summary>
    /// Write code to a module
    /// </summary>
    public void WriteModule(string filePath, string moduleName, string code)
    {
        // TODO: Implement
        throw new NotImplementedException("Access support coming soon");
    }
}
