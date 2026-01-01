namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when an invalid value is provided for a control property
/// </summary>
public class InvalidPropertyValueException : VbaMcpException
{
    public string PropertyName { get; }
    public object? ProvidedValue { get; }

    public InvalidPropertyValueException(string propertyName, object? value, string? filePath = null)
        : base($"Invalid value '{value}' for property '{propertyName}'", "INVALID_VALUE", filePath)
    {
        PropertyName = propertyName;
        ProvidedValue = value;
    }

    public InvalidPropertyValueException(string propertyName, object? value, Exception innerException, string? filePath = null)
        : base($"Invalid value '{value}' for property '{propertyName}'", "INVALID_VALUE", innerException, filePath)
    {
        PropertyName = propertyName;
        ProvidedValue = value;
    }
}
