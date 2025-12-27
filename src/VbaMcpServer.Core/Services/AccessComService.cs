using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Access;
using VbaMcpServer.Exceptions;
using VbaMcpServer.Helpers;
using VbaMcpServer.Models;

namespace VbaMcpServer.Services;

/// <summary>
/// Service for interacting with Access VBA projects via COM
/// </summary>
public class AccessComService
{
    private readonly ILogger<AccessComService> _logger;

    public AccessComService(ILogger<AccessComService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Check if Access is available
    /// </summary>
    public bool IsAccessAvailable()
    {
        try
        {
            var app = GetAccessApplication();
            return app != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Get the running Access application instance
    /// </summary>
    public Microsoft.Office.Interop.Access.Application? GetAccessApplication()
    {
        try
        {
            var app = ComHelper.GetActiveObject("Access.Application") as Microsoft.Office.Interop.Access.Application;
            if (app == null)
            {
                _logger.LogWarning("Access.Application COM object returned null");
                return null;
            }

            _logger.LogDebug("Successfully connected to Access application");
            return app;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to get Access application instance");
            return null;
        }
    }

    /// <summary>
    /// List all open Access databases
    /// </summary>
    public List<string> ListOpenDatabases()
    {
        var databases = new List<string>();
        var app = GetAccessApplication();

        if (app == null)
        {
            _logger.LogWarning("Access is not running");
            return databases;
        }

        try
        {
            // In Access, only one database can be open at a time
            if (app.CurrentProject != null && !string.IsNullOrEmpty(app.CurrentProject.FullName))
            {
                databases.Add(app.CurrentProject.FullName);
                _logger.LogDebug("Found open database: {Path}", app.CurrentProject.FullName);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing Access databases");
        }

        return databases;
    }

    /// <summary>
    /// Get the current database
    /// </summary>
    private Microsoft.Office.Interop.Access.Application? GetDatabase(string filePath)
    {
        var app = GetAccessApplication();
        if (app == null)
        {
            throw new ApplicationNotRunningException("Access");
        }

        // Check if the requested database is currently open
        var currentDbPath = app.CurrentProject?.FullName;
        if (string.IsNullOrEmpty(currentDbPath))
        {
            throw new FileNotFoundException($"No database is currently open in Access");
        }

        var normalizedCurrentPath = Path.GetFullPath(currentDbPath).ToLowerInvariant();
        var normalizedRequestPath = Path.GetFullPath(filePath).ToLowerInvariant();

        if (normalizedCurrentPath != normalizedRequestPath)
        {
            throw new FileNotFoundException(
                $"Database '{filePath}' is not open. Currently open: {currentDbPath}");
        }

        return app;
    }

    /// <summary>
    /// List all modules in a database
    /// </summary>
    public List<ModuleInfo> ListModules(string filePath)
    {
        var app = GetDatabase(filePath);
        var modules = new List<ModuleInfo>();

        try
        {
            // Access VBA modules through VBE
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            foreach (var component in vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>())
            {
                try
                {
                    var codeModule = component.CodeModule;
                    var lineCount = codeModule.CountOfLines;
                    var procedureCount = CountProcedures(codeModule);

                    var moduleInfo = new ModuleInfo
                    {
                        Name = component.Name,
                        Type = GetModuleTypeName(component.Type),
                        LineCount = lineCount,
                        ProcedureCount = procedureCount
                    };

                    modules.Add(moduleInfo);
                    _logger.LogDebug("Found module: {Name} ({Type}), {Lines} lines, {Procs} procedures",
                        moduleInfo.Name, moduleInfo.Type, lineCount, procedureCount);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read module info for {Name}", component.Name);
                }
            }
        }
        catch (UnauthorizedAccessException)
        {
            throw new VbaProjectAccessDeniedException(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing modules in {Path}", filePath);
            throw new VbaOperationException($"Failed to list modules: {ex.Message}", ex);
        }

        return modules;
    }

    /// <summary>
    /// Read the code from a module
    /// </summary>
    public string ReadModule(string filePath, string moduleName)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            var codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                _logger.LogDebug("Module {Module} is empty", moduleName);
                return string.Empty;
            }

            var code = codeModule.Lines[1, lineCount];
            _logger.LogDebug("Read {Lines} lines from module {Module}", lineCount, moduleName);

            return code;
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
            _logger.LogError(ex, "Error reading module {Module} from {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to read module '{moduleName}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Write code to a module
    /// </summary>
    public void WriteModule(string filePath, string moduleName, string code)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            var codeModule = component.CodeModule;

            // Delete existing code
            var existingLineCount = codeModule.CountOfLines;
            if (existingLineCount > 0)
            {
                codeModule.DeleteLines(1, existingLineCount);
            }

            // Insert new code
            if (!string.IsNullOrEmpty(code))
            {
                codeModule.InsertLines(1, code);
            }

            var newLineCount = code.Split('\n').Length;
            _logger.LogInformation("Wrote {Lines} lines to module {Module} in {Path}",
                newLineCount, moduleName, filePath);
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
            _logger.LogError(ex, "Error writing to module {Module} in {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to write to module '{moduleName}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Create a new module
    /// </summary>
    public void CreateModule(string filePath, string moduleName, VbaModuleType moduleType)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            // Check if module already exists
            var existing = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                throw new ModuleAlreadyExistsException(filePath, moduleName);
            }

            // Create new component
            var vbComponentType = ConvertToVbComponentType(moduleType);
            var newComponent = vbProject.VBComponents.Add(vbComponentType);
            newComponent.Name = moduleName;

            _logger.LogInformation("Created new {Type} module '{Module}' in {Path}",
                moduleType, moduleName, filePath);
        }
        catch (VbaProjectAccessDeniedException)
        {
            throw;
        }
        catch (ModuleAlreadyExistsException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error creating module {Module} in {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to create module '{moduleName}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Delete a module
    /// </summary>
    public void DeleteModule(string filePath, string moduleName)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            // Cannot delete document modules (forms, reports with code-behind)
            if (component.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document)
            {
                throw new InvalidOperationException(
                    $"Cannot delete document module '{moduleName}'. Forms and reports with code-behind cannot be deleted via COM.");
            }

            vbProject.VBComponents.Remove(component);

            _logger.LogInformation("Deleted module '{Module}' from {Path}", moduleName, filePath);
        }
        catch (VbaProjectAccessDeniedException)
        {
            throw;
        }
        catch (ModuleNotFoundException)
        {
            throw;
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting module {Module} from {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to delete module '{moduleName}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Export a module to a file
    /// </summary>
    public void ExportModule(string filePath, string moduleName, string outputPath)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            component.Export(outputPath);

            _logger.LogInformation("Exported module '{Module}' to {Output}", moduleName, outputPath);
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
            _logger.LogError(ex, "Error exporting module {Module} from {Path}", moduleName, filePath);
            throw new VbaOperationException($"Failed to export module '{moduleName}': {ex.Message}", ex);
        }
    }

    private string GetModuleTypeName(Microsoft.Vbe.Interop.vbext_ComponentType type)
    {
        return type switch
        {
            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule => "Standard Module",
            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule => "Class Module",
            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm => "UserForm",
            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document => "Document Module",
            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner => "ActiveX Designer",
            _ => "Unknown"
        };
    }

    private Microsoft.Vbe.Interop.vbext_ComponentType ConvertToVbComponentType(VbaModuleType moduleType)
    {
        return moduleType switch
        {
            VbaModuleType.StdModule => Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule,
            VbaModuleType.ClassModule => Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule,
            VbaModuleType.MSForm => Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm,
            _ => Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule
        };
    }

    private int CountProcedures(Microsoft.Vbe.Interop.CodeModule codeModule)
    {
        var procedureCount = 0;
        var lineCount = codeModule.CountOfLines;

        if (lineCount == 0) return 0;

        var processedProcedures = new HashSet<string>();

        for (int line = 1; line <= lineCount; line++)
        {
            try
            {
                var procName = codeModule.get_ProcOfLine(line, out var procKind);
                if (!string.IsNullOrEmpty(procName) && !processedProcedures.Contains(procName))
                {
                    processedProcedures.Add(procName);
                    procedureCount++;
                }
            }
            catch
            {
                // Ignore lines that are not in a procedure
            }
        }

        return procedureCount;
    }

    /// <summary>
    /// List all procedures in a module with detailed metadata
    /// </summary>
    public List<ProcedureInfo> ListProcedures(string filePath, string moduleName)
    {
        var app = GetDatabase(filePath);
        var procedures = new List<ProcedureInfo>();

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            var codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                _logger.LogDebug("Module {Module} is empty, no procedures to list", moduleName);
                return procedures;
            }

            var processedProcedures = new HashSet<string>();

            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out Microsoft.Vbe.Interop.vbext_ProcKind procKind);
                    if (!string.IsNullOrEmpty(procName) && !processedProcedures.Contains(procName))
                    {
                        processedProcedures.Add(procName);

                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);
                        var procType = GetProcedureTypeName(procKind);

                        // Extract access modifier from first line
                        var firstLine = codeModule.Lines[startLine, 1].Trim();
                        var accessModifier = GetAccessModifier(firstLine);

                        var procedureInfo = new ProcedureInfo
                        {
                            Name = procName,
                            Type = procType,
                            StartLine = startLine,
                            LineCount = procLineCount,
                            AccessModifier = accessModifier
                        };

                        procedures.Add(procedureInfo);
                        _logger.LogDebug("Found procedure: {Name} ({Type}), line {Start}, {Lines} lines, {Access}",
                            procName, procType, startLine, procLineCount, accessModifier);
                    }
                }
                catch
                {
                    // Ignore lines that are not in a procedure
                }
            }

            _logger.LogInformation("Listed {Count} procedures in module {Module} from {Path}",
                procedures.Count, moduleName, filePath);
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

        return procedures;
    }

    /// <summary>
    /// Read code of a specific procedure
    /// </summary>
    public string ReadProcedure(string filePath, string moduleName, string procedureName)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            var codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                throw new ArgumentException($"Module '{moduleName}' is empty, procedure '{procedureName}' not found", nameof(procedureName));
            }

            // Search for the procedure
            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out Microsoft.Vbe.Interop.vbext_ProcKind procKind);
                    if (!string.IsNullOrEmpty(procName) && procName.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
                    {
                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);
                        var code = codeModule.Lines[startLine, procLineCount];

                        _logger.LogDebug("Read procedure {Procedure} from module {Module}: {Lines} lines",
                            procedureName, moduleName, procLineCount);

                        return code;
                    }
                }
                catch
                {
                    // Continue searching
                }
            }

            throw new ArgumentException($"Procedure '{procedureName}' not found in module '{moduleName}'", nameof(procedureName));
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
    }

    /// <summary>
    /// Write/replace a specific procedure in a module
    /// </summary>
    public void WriteProcedure(string filePath, string moduleName, string procedureName, string newCode)
    {
        var app = GetDatabase(filePath);

        try
        {
            var vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            var component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            var codeModule = component.CodeModule;
            var lineCount = codeModule.CountOfLines;

            if (lineCount == 0)
            {
                throw new ArgumentException($"Module '{moduleName}' is empty, procedure '{procedureName}' not found", nameof(procedureName));
            }

            // Search for the procedure
            for (int line = 1; line <= lineCount; line++)
            {
                try
                {
                    var procName = codeModule.get_ProcOfLine(line, out Microsoft.Vbe.Interop.vbext_ProcKind procKind);
                    if (!string.IsNullOrEmpty(procName) && procName.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
                    {
                        var startLine = codeModule.get_ProcStartLine(procName, procKind);
                        var procLineCount = codeModule.get_ProcCountLines(procName, procKind);

                        // Delete existing procedure
                        codeModule.DeleteLines(startLine, procLineCount);

                        // Insert new code at the same position
                        if (!string.IsNullOrEmpty(newCode))
                        {
                            codeModule.InsertLines(startLine, newCode);
                        }

                        var newLineCount = string.IsNullOrEmpty(newCode) ? 0 : newCode.Split('\n').Length;
                        _logger.LogInformation("Wrote procedure {Procedure} to module {Module} in {Path}: {Lines} lines",
                            procedureName, moduleName, filePath, newLineCount);

                        return;
                    }
                }
                catch
                {
                    // Continue searching
                }
            }

            throw new ArgumentException($"Procedure '{procedureName}' not found in module '{moduleName}'", nameof(procedureName));
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
    }

    private string GetProcedureTypeName(Microsoft.Vbe.Interop.vbext_ProcKind procKind)
    {
        return procKind switch
        {
            Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc => "Sub/Function",
            Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Get => "Property Get",
            Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Let => "Property Let",
            Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Set => "Property Set",
            _ => "Unknown"
        };
    }

    private string? GetAccessModifier(string firstLine)
    {
        var lowerLine = firstLine.ToLowerInvariant();
        if (lowerLine.StartsWith("public ")) return "Public";
        if (lowerLine.StartsWith("private ")) return "Private";
        if (lowerLine.StartsWith("friend ")) return "Friend";
        return "Public"; // Default in VBA if not specified
    }
}
