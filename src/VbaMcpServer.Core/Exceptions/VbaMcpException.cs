namespace VbaMcpServer.Exceptions;

/// <summary>
/// Base exception class for VBA MCP Server errors
/// </summary>
public class VbaMcpException : Exception
{
    /// <summary>
    /// Error code for categorizing the exception
    /// </summary>
    public string ErrorCode { get; }

    /// <summary>
    /// File path related to the error (if applicable)
    /// </summary>
    public string? FilePath { get; }

    /// <summary>
    /// Module name related to the error (if applicable)
    /// </summary>
    public string? ModuleName { get; }

    public VbaMcpException(string message, string errorCode,
                          string? filePath = null, string? moduleName = null)
        : base(message)
    {
        ErrorCode = errorCode;
        FilePath = filePath;
        ModuleName = moduleName;
    }

    public VbaMcpException(string message, string errorCode, Exception innerException,
                          string? filePath = null, string? moduleName = null)
        : base(message, innerException)
    {
        ErrorCode = errorCode;
        FilePath = filePath;
        ModuleName = moduleName;
    }
}
