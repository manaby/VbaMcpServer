namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a specified form is not found in the Access database
/// </summary>
public class FormNotFoundException : VbaMcpException
{
    public FormNotFoundException(string formName, string? filePath = null)
        : base($"Form not found: {formName}", "FORM_NOT_FOUND", filePath)
    {
    }

    public FormNotFoundException(string formName, Exception innerException, string? filePath = null)
        : base($"Form not found: {formName}", "FORM_NOT_FOUND", innerException, filePath)
    {
    }
}
