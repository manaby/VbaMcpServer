namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when attempting to create a module that already exists
/// </summary>
public class ModuleAlreadyExistsException : VbaMcpException
{
    public ModuleAlreadyExistsException(string filePath, string moduleName)
        : base($"Module '{moduleName}' already exists in '{filePath}'.",
               "MODULE_ALREADY_EXISTS", filePath, moduleName)
    {
    }
}
