namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a specified control is not found in a form or report
/// </summary>
public class ControlNotFoundException : VbaMcpException
{
    public string ObjectName { get; }

    public ControlNotFoundException(string controlName, string objectName, string? filePath = null)
        : base($"Control '{controlName}' not found in {objectName}", "CONTROL_NOT_FOUND", filePath)
    {
        ObjectName = objectName;
    }

    public ControlNotFoundException(string controlName, string objectName, Exception innerException, string? filePath = null)
        : base($"Control '{controlName}' not found in {objectName}", "CONTROL_NOT_FOUND", innerException, filePath)
    {
        ObjectName = objectName;
    }
}
