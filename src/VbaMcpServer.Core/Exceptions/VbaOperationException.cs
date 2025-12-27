namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a VBA operation fails
/// </summary>
public class VbaOperationException : VbaMcpException
{
    public VbaOperationException(string message, Exception? innerException = null, string? filePath = null, string? moduleName = null)
        : base(message, "VBA_OPERATION_FAILED", innerException!, filePath, moduleName)
    {
    }
}
