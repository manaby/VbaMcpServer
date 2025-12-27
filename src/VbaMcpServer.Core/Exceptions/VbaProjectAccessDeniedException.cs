namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when VBA project access is denied due to Trust Center settings
/// </summary>
public class VbaProjectAccessDeniedException : VbaMcpException
{
    public VbaProjectAccessDeniedException(string filePath)
        : base($"VBA project access denied for '{filePath}'. Please enable 'Trust access to the VBA project object model' in Office Trust Center settings.",
               "VBA_PROJECT_ACCESS_DENIED", filePath)
    {
    }
}
