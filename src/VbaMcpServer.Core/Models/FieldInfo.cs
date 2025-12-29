namespace VbaMcpServer.Models;

/// <summary>
/// Information about a table field/column
/// </summary>
public class FieldInfo
{
    /// <summary>
    /// Field name
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Data type (e.g., "Text", "Number (Long)", "Date/Time")
    /// </summary>
    public required string DataType { get; init; }

    /// <summary>
    /// Field size
    /// </summary>
    public int Size { get; init; }

    /// <summary>
    /// Whether the field is required (cannot be null)
    /// </summary>
    public bool Required { get; init; }

    /// <summary>
    /// Whether zero-length strings are allowed (for text fields)
    /// </summary>
    public bool AllowZeroLength { get; init; }

    /// <summary>
    /// Default value for the field
    /// </summary>
    public string? DefaultValue { get; init; }

    /// <summary>
    /// Validation rule expression
    /// </summary>
    public string? ValidationRule { get; init; }

    /// <summary>
    /// Whether this field is a primary key
    /// </summary>
    public bool IsPrimaryKey { get; init; }

    /// <summary>
    /// Whether this field is indexed
    /// </summary>
    public bool IsIndexed { get; init; }
}
