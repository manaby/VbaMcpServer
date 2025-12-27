namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a VBA module is not found
/// </summary>
public class ModuleNotFoundException : VbaMcpException
{
    public ModuleNotFoundException(string moduleName, string? filePath = null)
        : base($"Module not found: {moduleName}", "MODULE_NOT_FOUND", filePath, moduleName)
    {
    }

    public ModuleNotFoundException(string moduleName, Exception innerException, string? filePath = null)
        : base($"Module not found: {moduleName}", "MODULE_NOT_FOUND", innerException, filePath, moduleName)
    {
    }
}
