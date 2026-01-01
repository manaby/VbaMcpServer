namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a specified property is not found on a control
/// </summary>
public class PropertyNotFoundException : VbaMcpException
{
    public string PropertyName { get; }
    public string ControlName { get; }

    public PropertyNotFoundException(string propertyName, string controlName, string? filePath = null)
        : base($"Property '{propertyName}' not found on control '{controlName}'", "PROPERTY_NOT_FOUND", filePath)
    {
        PropertyName = propertyName;
        ControlName = controlName;
    }

    public PropertyNotFoundException(string propertyName, string controlName, Exception innerException, string? filePath = null)
        : base($"Property '{propertyName}' not found on control '{controlName}'", "PROPERTY_NOT_FOUND", innerException, filePath)
    {
        PropertyName = propertyName;
        ControlName = controlName;
    }
}
