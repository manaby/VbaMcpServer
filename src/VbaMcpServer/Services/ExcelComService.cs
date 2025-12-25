using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using VbaMcpServer.Helpers;
using VbaMcpServer.Models;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;

namespace VbaMcpServer.Services;

/// <summary>
/// Service for interacting with Excel VBA projects via COM
/// </summary>
public class ExcelComService
{
    private readonly ILogger<ExcelComService> _logger;

    public ExcelComService(ILogger<ExcelComService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Check if Excel is available
    /// </summary>
    public bool IsExcelAvailable()
    {
        try
        {
            var excel = GetExcelApplication();
            return excel != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Get the running Excel application instance
    /// </summary>
    private Excel.Application? GetExcelApplication()
    {
        try
        {
            // Use ComHelper for .NET 8+ compatibility (Marshal.GetActiveObject is not supported)
            return (Excel.Application)ComHelper.GetActiveObject("Excel.Application");
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            // MK_E_UNAVAILABLE - Excel is not running
            _logger.LogDebug("Excel is not running");
            return null;
        }
    }

    /// <summary>
    /// Get a workbook by file path
    /// </summary>
    public Excel.Workbook? GetWorkbook(string filePath)
    {
        var excel = GetExcelApplication();
        if (excel == null) return null;

        var normalizedPath = Path.GetFullPath(filePath);
        
        foreach (Excel.Workbook wb in excel.Workbooks)
        {
            if (string.Equals(wb.FullName, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                return wb;
            }
        }

        _logger.LogWarning("Workbook not found: {Path}", filePath);
        return null;
    }

    /// <summary>
    /// List all open workbooks
    /// </summary>
    public List<string> ListOpenWorkbooks()
    {
        var result = new List<string>();
        var excel = GetExcelApplication();
        
        if (excel == null) return result;

        foreach (Excel.Workbook wb in excel.Workbooks)
        {
            result.Add(wb.FullName);
        }

        return result;
    }

    /// <summary>
    /// List all modules in a workbook
    /// </summary>
    public List<ModuleInfo> ListModules(string filePath)
    {
        var result = new List<ModuleInfo>();
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            
            foreach (VBIDE.VBComponent component in vbProject.VBComponents)
            {
                var moduleInfo = new ModuleInfo
                {
                    Name = component.Name,
                    Type = GetModuleTypeName(component.Type),
                    LineCount = component.CodeModule.CountOfLines,
                    ProcedureCount = CountProcedures(component.CodeModule)
                };
                result.Add(moduleInfo);
            }
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }

        return result;
    }

    /// <summary>
    /// Read the code from a module
    /// </summary>
    public string ReadModule(string filePath, string moduleName)
    {
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            var component = vbProject.VBComponents.Item(moduleName);
            var codeModule = component.CodeModule;

            if (codeModule.CountOfLines == 0)
            {
                return string.Empty;
            }

            return codeModule.Lines[1, codeModule.CountOfLines];
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }
        catch (COMException ex) when (ex.Message.Contains("Subscript out of range"))
        {
            throw new ArgumentException($"Module not found: {moduleName}", ex);
        }
    }

    /// <summary>
    /// Write code to a module (replaces entire content)
    /// </summary>
    public void WriteModule(string filePath, string moduleName, string code)
    {
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            var component = vbProject.VBComponents.Item(moduleName);
            var codeModule = component.CodeModule;

            // Delete existing code
            if (codeModule.CountOfLines > 0)
            {
                codeModule.DeleteLines(1, codeModule.CountOfLines);
            }

            // Insert new code
            if (!string.IsNullOrEmpty(code))
            {
                codeModule.InsertLines(1, code);
            }

            _logger.LogInformation("Module {Module} updated in {File}", moduleName, filePath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }
        catch (COMException ex) when (ex.Message.Contains("Subscript out of range"))
        {
            throw new ArgumentException($"Module not found: {moduleName}", ex);
        }
    }

    /// <summary>
    /// Create a new module
    /// </summary>
    public void CreateModule(string filePath, string moduleName, VbaModuleType moduleType)
    {
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            var componentType = (VBIDE.vbext_ComponentType)moduleType;
            var component = vbProject.VBComponents.Add(componentType);
            component.Name = moduleName;

            _logger.LogInformation("Module {Module} created in {File}", moduleName, filePath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }
    }

    /// <summary>
    /// Delete a module
    /// </summary>
    public void DeleteModule(string filePath, string moduleName)
    {
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            var component = vbProject.VBComponents.Item(moduleName);
            
            // Cannot delete document modules
            if (component.Type == VBIDE.vbext_ComponentType.vbext_ct_Document)
            {
                throw new InvalidOperationException($"Cannot delete document module: {moduleName}");
            }

            vbProject.VBComponents.Remove(component);

            _logger.LogInformation("Module {Module} deleted from {File}", moduleName, filePath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }
        catch (COMException ex) when (ex.Message.Contains("Subscript out of range"))
        {
            throw new ArgumentException($"Module not found: {moduleName}", ex);
        }
    }

    /// <summary>
    /// Export a module to a file
    /// </summary>
    public void ExportModule(string filePath, string moduleName, string outputPath)
    {
        var workbook = GetWorkbook(filePath);
        
        if (workbook == null)
        {
            throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
        }

        try
        {
            var vbProject = workbook.VBProject;
            var component = vbProject.VBComponents.Item(moduleName);
            component.Export(outputPath);

            _logger.LogInformation("Module {Module} exported to {Output}", moduleName, outputPath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw new UnauthorizedAccessException(
                "VBA project access is not trusted. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings.",
                ex);
        }
        catch (COMException ex) when (ex.Message.Contains("Subscript out of range"))
        {
            throw new ArgumentException($"Module not found: {moduleName}", ex);
        }
    }

    private static string GetModuleTypeName(VBIDE.vbext_ComponentType type)
    {
        return type switch
        {
            VBIDE.vbext_ComponentType.vbext_ct_StdModule => "StdModule",
            VBIDE.vbext_ComponentType.vbext_ct_ClassModule => "ClassModule",
            VBIDE.vbext_ComponentType.vbext_ct_MSForm => "UserForm",
            VBIDE.vbext_ComponentType.vbext_ct_Document => "Document",
            _ => "Unknown"
        };
    }

    private static int CountProcedures(VBIDE.CodeModule codeModule)
    {
        if (codeModule.CountOfLines == 0) return 0;

        int count = 0;
        int line = 1;

        while (line <= codeModule.CountOfLines)
        {
            var procName = codeModule.ProcOfLine[line, out VBIDE.vbext_ProcKind _];
            if (!string.IsNullOrEmpty(procName))
            {
                count++;
                var procLines = codeModule.ProcCountLines[procName, VBIDE.vbext_ProcKind.vbext_pk_Proc];
                line += procLines;
            }
            else
            {
                line++;
            }
        }

        return count;
    }
}
