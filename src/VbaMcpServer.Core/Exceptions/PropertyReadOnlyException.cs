namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when attempting to set a read-only property on a control
/// </summary>
public class PropertyReadOnlyException : VbaMcpException
{
    public string PropertyName { get; }
    public string ControlName { get; }

    public PropertyReadOnlyException(string propertyName, string controlName, string? filePath = null)
        : base($"Property '{propertyName}' on control '{controlName}' is read-only", "PROPERTY_READ_ONLY", filePath)
    {
        PropertyName = propertyName;
        ControlName = controlName;
    }

    public PropertyReadOnlyException(string propertyName, string controlName, Exception innerException, string? filePath = null)
        : base($"Property '{propertyName}' on control '{controlName}' is read-only", "PROPERTY_READ_ONLY", innerException, filePath)
    {
        PropertyName = propertyName;
        ControlName = controlName;
    }
}
