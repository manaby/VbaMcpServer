namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a COM operation fails
/// </summary>
public class ComOperationException : VbaMcpException
{
    /// <summary>
    /// COM HRESULT error code
    /// </summary>
    public new int HResult { get; }

    public ComOperationException(string message, int hresult,
                                string? filePath = null, string? moduleName = null)
        : base(message, "COM_ERROR", filePath, moduleName)
    {
        HResult = hresult;
    }

    public ComOperationException(string message, int hresult, Exception innerException,
                                string? filePath = null, string? moduleName = null)
        : base(message, "COM_ERROR", innerException, filePath, moduleName)
    {
        HResult = hresult;
    }
}
