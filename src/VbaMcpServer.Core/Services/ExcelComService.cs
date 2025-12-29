using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using VbaMcpServer.Exceptions;
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
        Excel.Application? excel = null;
        try
        {
            excel = GetExcelApplication();
            return excel != null;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(excel);
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
        catch (COMException ex) when (ComErrorCodes.IsApplicationUnavailable(ex.HResult))
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
        Excel.Application? excel = null;
        Excel.Workbook? wb = null;
        Excel.Workbook? result = null;

        try
        {
            excel = GetExcelApplication();
            if (excel == null) return null;

            var normalizedPath = Path.GetFullPath(filePath);

            foreach (Excel.Workbook workbook in excel.Workbooks)
            {
                wb = workbook;
                try
                {
                    if (string.Equals(wb.FullName, normalizedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        result = wb;
                        wb = null; // Don't release the workbook we're returning
                        return result;
                    }
                }
                finally
                {
                    // Release workbook if it's not the one we're returning
                    if (wb != null && wb != result)
                    {
                        ReleaseComObject(wb);
                        wb = null;
                    }
                }
            }

            _logger.LogWarning("Workbook not found: {Path}", filePath);
            return null;
        }
        finally
        {
            ReleaseComObject(excel);
        }
    }

    /// <summary>
    /// List all open workbooks
    /// </summary>
    public List<string> ListOpenWorkbooks()
    {
        var result = new List<string>();
        Excel.Application? excel = null;
        Excel.Workbook? wb = null;

        try
        {
            excel = GetExcelApplication();
            if (excel == null) return result;

            foreach (Excel.Workbook workbook in excel.Workbooks)
            {
                wb = workbook;
                try
                {
                    result.Add(wb.FullName);
                }
                finally
                {
                    ReleaseComObject(wb);
                    wb = null;
                }
            }

            return result;
        }
        finally
        {
            ReleaseComObject(excel);
        }
    }

    /// <summary>
    /// List all modules in a workbook
    /// </summary>
    public List<ModuleInfo> ListModules(string filePath)
    {
        var result = new List<ModuleInfo>();
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;

            foreach (VBIDE.VBComponent comp in vbProject.VBComponents)
            {
                component = comp;
                try
                {
                    codeModule = component.CodeModule;
                    var moduleInfo = new ModuleInfo
                    {
                        Name = component.Name,
                        Type = GetModuleTypeName(component.Type),
                        LineCount = codeModule.CountOfLines,
                        ProcedureCount = CountProcedures(codeModule)
                    };
                    result.Add(moduleInfo);
                }
                finally
                {
                    ReleaseComObject(codeModule);
                    ReleaseComObject(component);
                    codeModule = null;
                    component = null;
                }
            }
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        finally
        {
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }

        return result;
    }

    /// <summary>
    /// Read the code from a module
    /// </summary>
    public string ReadModule(string filePath, string moduleName)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;
            component = vbProject.VBComponents.Item(moduleName);
            codeModule = component.CodeModule;

            if (codeModule.CountOfLines == 0)
            {
                return string.Empty;
            }

            return codeModule.Lines[1, codeModule.CountOfLines];
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        catch (COMException ex) when (ComErrorCodes.IsNotFoundError(ex.HResult) || ex.Message.Contains("Subscript out of range"))
        {
            throw new ModuleNotFoundException(moduleName, ex, filePath);
        }
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Write code to a module (replaces entire content)
    /// </summary>
    public void WriteModule(string filePath, string moduleName, string code)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;
            component = vbProject.VBComponents.Item(moduleName);
            codeModule = component.CodeModule;

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
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        catch (COMException ex) when (ComErrorCodes.IsNotFoundError(ex.HResult) || ex.Message.Contains("Subscript out of range"))
        {
            throw new ModuleNotFoundException(moduleName, ex, filePath);
        }
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Create a new module
    /// </summary>
    public void CreateModule(string filePath, string moduleName, VbaModuleType moduleType)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;
            var componentType = (VBIDE.vbext_ComponentType)moduleType;
            component = vbProject.VBComponents.Add(componentType);
            component.Name = moduleName;

            _logger.LogInformation("Module {Module} created in {File}", moduleName, filePath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        finally
        {
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Delete a module
    /// </summary>
    public void DeleteModule(string filePath, string moduleName)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;
            component = vbProject.VBComponents.Item(moduleName);

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
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        catch (COMException ex) when (ComErrorCodes.IsNotFoundError(ex.HResult) || ex.Message.Contains("Subscript out of range"))
        {
            throw new ModuleNotFoundException(moduleName, ex, filePath);
        }
        finally
        {
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Export a module to a file
    /// </summary>
    public void ExportModule(string filePath, string moduleName, string outputPath)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new FileNotFoundException($"Workbook not found or not open: {filePath}");
            }

            vbProject = workbook.VBProject;
            component = vbProject.VBComponents.Item(moduleName);
            component.Export(outputPath);

            _logger.LogInformation("Module {Module} exported to {Output}", moduleName, outputPath);
        }
        catch (COMException ex) when (ex.Message.Contains("programmatic access"))
        {
            throw VbaAccessException.CreateTrustCenterError(filePath);
        }
        catch (COMException ex) when (ComErrorCodes.IsNotFoundError(ex.HResult) || ex.Message.Contains("Subscript out of range"))
        {
            throw new ModuleNotFoundException(moduleName, ex, filePath);
        }
        finally
        {
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
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

    /// <summary>
    /// List all procedures in a module with detailed metadata
    /// </summary>
    public List<ProcedureInfo> ListProcedures(string filePath, string moduleName)
    {
        var procedures = new List<ProcedureInfo>();
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            vbProject = workbook.VBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<VBIDE.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                return procedures;
            }

            var processedProcedures = new HashSet<string>();

            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out VBIDE.vbext_ProcKind procKind);

                    if (!string.IsNullOrEmpty(procName) && !processedProcedures.Contains(procName))
                    {
                        processedProcedures.Add(procName);

                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);

                        // Get access modifier and procedure type by analyzing the first line of the procedure
                        string? accessModifier = null;
                        string procType;
                        if (startLine > 0 && startLine <= lineCount)
                        {
                            var firstLine = codeModule.Lines[startLine, 1].Trim();
                            accessModifier = GetAccessModifier(firstLine);
                            procType = GetProcedureTypeName(procKind, firstLine);
                        }
                        else
                        {
                            procType = GetProcedureTypeName(procKind);
                        }

                        var procedureInfo = new ProcedureInfo
                        {
                            Name = procName,
                            Type = procType,
                            StartLine = startLine,
                            LineCount = procLineCount,
                            AccessModifier = accessModifier
                        };

                        procedures.Add(procedureInfo);
                        _logger.LogDebug("Found procedure: {Name} ({Type}), lines {Start}-{End}",
                            procName, procType, startLine, startLine + procLineCount - 1);
                    }
                }
                catch
                {
                    // Skip lines that are not in a procedure
                }
            }

            _logger.LogDebug("Found {Count} procedures in module {Module}", procedures.Count, moduleName);
        }
        catch (VbaProjectAccessDeniedException)
        {
            throw;
        }
        catch (ModuleNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing procedures in module {Module} from {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to list procedures in module '{moduleName}': {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }

        return procedures;
    }

    /// <summary>
    /// Read code of a specific procedure
    /// </summary>
    public string ReadProcedure(string filePath, string moduleName, string procedureName)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            vbProject = workbook.VBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<VBIDE.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                throw new ArgumentException($"Module '{moduleName}' is empty");
            }

            // Search for the procedure
            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out VBIDE.vbext_ProcKind procKind);

                    if (!string.IsNullOrEmpty(procName) && procName.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
                    {
                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);

                        var code = codeModule.Lines[startLine, procLineCount];
                        _logger.LogDebug("Read procedure {Procedure} ({Lines} lines) from module {Module}",
                            procedureName, procLineCount, moduleName);

                        return code;
                    }
                }
                catch
                {
                    // Continue searching
                }
            }

            throw new ArgumentException($"Procedure '{procedureName}' not found in module '{moduleName}'");
        }
        catch (VbaProjectAccessDeniedException)
        {
            throw;
        }
        catch (ModuleNotFoundException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error reading procedure {Procedure} from module {Module} in {Path}",
                procedureName, moduleName, filePath);
            throw new VbaOperationException($"Failed to read procedure '{procedureName}': {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Write/replace a specific procedure in a module
    /// </summary>
    public void WriteProcedure(string filePath, string moduleName, string procedureName, string newCode)
    {
        Excel.Workbook? workbook = null;
        VBIDE.VBProject? vbProject = null;
        VBIDE.VBComponent? component = null;
        VBIDE.CodeModule? codeModule = null;

        try
        {
            workbook = GetWorkbook(filePath);

            if (workbook == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            vbProject = workbook.VBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<VBIDE.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                throw new ArgumentException($"Module '{moduleName}' is empty");
            }

            // Search for the procedure
            bool found = false;
            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out VBIDE.vbext_ProcKind procKind);

                    if (!string.IsNullOrEmpty(procName) && procName.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
                    {
                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);

                        // Delete the existing procedure
                        codeModule.DeleteLines(startLine, procLineCount);

                        // Insert the new code at the same position
                        if (!string.IsNullOrEmpty(newCode))
                        {
                            codeModule.InsertLines(startLine, newCode);
                        }

                        found = true;
                        _logger.LogInformation("Replaced procedure {Procedure} in module {Module} in {Path}",
                            procedureName, moduleName, filePath);
                        break;
                    }
                }
                catch
                {
                    // Continue searching
                }
            }

            if (!found)
            {
                throw new ArgumentException($"Procedure '{procedureName}' not found in module '{moduleName}'");
            }
        }
        catch (VbaProjectAccessDeniedException)
        {
            throw;
        }
        catch (ModuleNotFoundException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error writing procedure {Procedure} to module {Module} in {Path}",
                procedureName, moduleName, filePath);
            throw new VbaOperationException($"Failed to write procedure '{procedureName}': {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(workbook);
        }
    }

    private string GetProcedureTypeName(VBIDE.vbext_ProcKind procKind, string? firstLine = null)
    {
        // For regular procedures (Sub/Function), parse the first line to determine the exact type
        if (procKind == VBIDE.vbext_ProcKind.vbext_pk_Proc && !string.IsNullOrWhiteSpace(firstLine))
        {
            // Normalize the line: add spaces at boundaries and convert to lowercase
            var normalized = " " + firstLine.ToLowerInvariant().Replace("\t", " ").Trim() + " ";

            if (normalized.Contains(" function "))
                return "Function";
            if (normalized.Contains(" sub "))
                return "Sub";
        }

        // Fallback to original logic
        return procKind switch
        {
            VBIDE.vbext_ProcKind.vbext_pk_Proc => "Sub/Function",
            VBIDE.vbext_ProcKind.vbext_pk_Get => "Property Get",
            VBIDE.vbext_ProcKind.vbext_pk_Let => "Property Let",
            VBIDE.vbext_ProcKind.vbext_pk_Set => "Property Set",
            _ => "Unknown"
        };
    }

    private string? GetAccessModifier(string firstLine)
    {
        var lowerLine = firstLine.ToLowerInvariant();

        if (lowerLine.StartsWith("public "))
            return "Public";
        if (lowerLine.StartsWith("private "))
            return "Private";
        if (lowerLine.StartsWith("friend "))
            return "Friend";

        // Default to Public if not specified
        return "Public";
    }

    /// <summary>
    /// Release COM object
    /// </summary>
    private void ReleaseComObject(object? comObject)
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to release COM object");
            }
        }
    }
}
