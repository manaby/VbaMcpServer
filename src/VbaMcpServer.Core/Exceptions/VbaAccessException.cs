namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when VBA project access is denied
/// </summary>
public class VbaAccessException : VbaMcpException
{
    public VbaAccessException(string message, string? filePath = null, string? moduleName = null)
        : base(message, "VBA_ACCESS_DENIED", filePath, moduleName)
    {
    }

    public VbaAccessException(string message, Exception innerException,
                            string? filePath = null, string? moduleName = null)
        : base(message, "VBA_ACCESS_DENIED", innerException, filePath, moduleName)
    {
    }

    /// <summary>
    /// Creates a standard VBA access denied exception with helpful message
    /// </summary>
    public static VbaAccessException CreateTrustCenterError(string? filePath = null)
    {
        return new VbaAccessException(
            "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
            filePath);
    }
}
