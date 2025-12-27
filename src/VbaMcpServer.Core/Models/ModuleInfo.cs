namespace VbaMcpServer.Models;

/// <summary>
/// Represents information about a VBA module
/// </summary>
public class ModuleInfo
{
    /// <summary>
    /// Name of the module
    /// </summary>
    public required string Name { get; set; }

    /// <summary>
    /// Type of the module (StdModule, ClassModule, MSForm, Document)
    /// </summary>
    public required string Type { get; set; }

    /// <summary>
    /// Number of lines in the module
    /// </summary>
    public int LineCount { get; set; }

    /// <summary>
    /// Number of procedures in the module
    /// </summary>
    public int ProcedureCount { get; set; }
}

/// <summary>
/// VBA module types
/// </summary>
public enum VbaModuleType
{
    StdModule = 1,
    ClassModule = 2,
    MSForm = 3,
    Document = 100
}
