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
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            // Access VBA modules through VBE
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            foreach (var comp in vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>())
            {
                component = comp;
                try
                {
                    codeModule = component.CodeModule;
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
                finally
                {
                    ReleaseComObject(codeModule);
                    ReleaseComObject(component);
                    codeModule = null;
                    component = null;
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
        finally
        {
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }

        return modules;
    }

    /// <summary>
    /// Read the code from a module
    /// </summary>
    public string ReadModule(string filePath, string moduleName)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
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
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Write code to a module
    /// </summary>
    public void WriteModule(string filePath, string moduleName, string code)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;

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
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Create a new module
    /// </summary>
    public void CreateModule(string filePath, string moduleName, VbaModuleType moduleType)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? existing = null;
        Microsoft.Vbe.Interop.VBComponent? newComponent = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            // Check if module already exists
            existing = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                throw new ModuleAlreadyExistsException(filePath, moduleName);
            }

            // Create new component
            var vbComponentType = ConvertToVbComponentType(moduleType);
            newComponent = vbProject.VBComponents.Add(vbComponentType);
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
        finally
        {
            ReleaseComObject(newComponent);
            ReleaseComObject(existing);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Delete a module
    /// </summary>
    public void DeleteModule(string filePath, string moduleName)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
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
        finally
        {
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Export a module to a file
    /// </summary>
    public void ExportModule(string filePath, string moduleName, string outputPath)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
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
        finally
        {
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
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
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
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

                        // Extract access modifier and procedure type from first line
                        var firstLine = codeModule.Lines[startLine, 1].Trim();
                        var accessModifier = GetAccessModifier(firstLine);
                        var procType = GetProcedureTypeName(procKind, firstLine);

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
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }

        return procedures;
    }

    /// <summary>
    /// Read code of a specific procedure
    /// </summary>
    public string ReadProcedure(string filePath, string moduleName, string procedureName)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
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
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Write/replace a specific procedure in a module
    /// </summary>
    public void WriteProcedure(string filePath, string moduleName, string procedureName, string newCode)
    {
        var app = GetDatabase(filePath);
        Microsoft.Vbe.Interop.VBProject? vbProject = null;
        Microsoft.Vbe.Interop.VBComponent? component = null;
        Microsoft.Vbe.Interop.CodeModule? codeModule = null;

        try
        {
            vbProject = app!.VBE.ActiveVBProject;
            if (vbProject == null)
            {
                throw new VbaProjectAccessDeniedException(filePath);
            }

            component = vbProject.VBComponents.Cast<Microsoft.Vbe.Interop.VBComponent>()
                .FirstOrDefault(c => c.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));

            if (component == null)
            {
                throw new ModuleNotFoundException(filePath, moduleName);
            }

            codeModule = component.CodeModule;
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
        finally
        {
            ReleaseComObject(codeModule);
            ReleaseComObject(component);
            ReleaseComObject(vbProject);
            ReleaseComObject(app);
        }
    }

    private string GetProcedureTypeName(Microsoft.Vbe.Interop.vbext_ProcKind procKind, string? firstLine = null)
    {
        // For regular procedures (Sub/Function), parse the first line to determine the exact type
        if (procKind == Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc && !string.IsNullOrWhiteSpace(firstLine))
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

    #region Table Operations

    /// <summary>
    /// List all tables in the database
    /// </summary>
    public List<TableInfo> ListTables(string filePath, bool includeSystemTables = false)
    {
        var app = GetDatabase(filePath);
        var tables = new List<TableInfo>();
        dynamic? currentDb = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            foreach (var tableDef in currentDb.TableDefs)
            {
                try
                {
                    var name = tableDef.Name;

                    // Skip system tables unless requested
                    if (!includeSystemTables)
                    {
                        if (name.StartsWith("MSys") || name.StartsWith("~"))
                            continue;

                        // Check table attributes for system table flag
                        int attributes = tableDef.Attributes;
                        const int dbSystemObject = unchecked((int)0x80000002);
                        if ((attributes & dbSystemObject) != 0)
                            continue;
                    }

                    var recordCount = 0;
                    try
                    {
                        recordCount = tableDef.RecordCount;
                    }
                    catch
                    {
                        // Some tables may not support RecordCount
                    }

                    var tableType = GetTableType(tableDef.Attributes);

                    var tableInfo = new TableInfo
                    {
                        Name = name,
                        Type = tableType,
                        RecordCount = recordCount,
                        DateCreated = TryGetDateTime(tableDef.DateCreated),
                        DateModified = TryGetDateTime(tableDef.LastUpdated)
                    };

                    tables.Add(tableInfo);
                    _logger.LogDebug("Found table: {Name} ({Type}), {Count} records",
                        (string)name, (string)tableType, recordCount);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read table info");
                }
            }

            _logger.LogInformation("Listed {Count} tables in {Path}", tables.Count, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing tables in {Path}", filePath);
            throw new VbaOperationException($"Failed to list tables: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(currentDb);
        }

        return tables;
    }

    /// <summary>
    /// Get the structure of a table
    /// </summary>
    public List<FieldInfo> GetTableStructure(string filePath, string tableName)
    {
        var app = GetDatabase(filePath);
        var fields = new List<FieldInfo>();
        dynamic? currentDb = null;
        dynamic? tableDef = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            try
            {
                tableDef = currentDb.TableDefs[tableName];
            }
            catch
            {
                throw new TableNotFoundException(tableName, filePath);
            }

            // Get primary key fields
            var primaryKeyFields = new HashSet<string>();
            try
            {
                foreach (var index in tableDef.Indexes)
                {
                    if (index.Primary)
                    {
                        foreach (var field in index.Fields)
                        {
                            primaryKeyFields.Add(field.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to read indexes for table {Table}", tableName);
            }

            // Get field information
            foreach (var field in tableDef.Fields)
            {
                try
                {
                    var fieldName = field.Name;
                    var dataType = MapDataType(field.Type);
                    var size = field.Size;
                    var required = field.Required;
                    var allowZeroLength = false;

                    try
                    {
                        allowZeroLength = field.AllowZeroLength;
                    }
                    catch
                    {
                        // Not all field types support this property
                    }

                    var defaultValue = field.DefaultValue?.ToString();
                    var validationRule = field.ValidationRule?.ToString();
                    var isPrimaryKey = primaryKeyFields.Contains(fieldName);

                    // Check if field is indexed
                    var isIndexed = false;
                    try
                    {
                        isIndexed = field.Attributes != 0;
                    }
                    catch
                    {
                        // Ignore
                    }

                    var fieldInfo = new FieldInfo
                    {
                        Name = fieldName,
                        DataType = dataType,
                        Size = size,
                        Required = required,
                        AllowZeroLength = allowZeroLength,
                        DefaultValue = defaultValue,
                        ValidationRule = validationRule,
                        IsPrimaryKey = isPrimaryKey,
                        IsIndexed = isIndexed
                    };

                    fields.Add(fieldInfo);
                    _logger.LogDebug("Found field: {Name} ({Type})", (string)fieldName, (string)dataType);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read field info");
                }
            }

            _logger.LogInformation("Retrieved structure for table {Table}: {Count} fields",
                tableName, fields.Count);
        }
        catch (TableNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting table structure for {Table} in {Path}",
                tableName, filePath);
            throw new VbaOperationException($"Failed to get table structure: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(tableDef);
            ReleaseComObject(currentDb);
        }

        return fields;
    }

    /// <summary>
    /// Get data from a table
    /// </summary>
    public TableDataResult GetTableData(string filePath, string tableName,
        string? whereClause = null, int limit = 100, int offset = 0)
    {
        if (limit <= 0 || limit > 1000)
        {
            throw new ArgumentException("Limit must be between 1 and 1000", nameof(limit));
        }

        if (offset < 0)
        {
            throw new ArgumentException("Offset must be non-negative", nameof(offset));
        }

        if (!string.IsNullOrWhiteSpace(whereClause))
        {
            ValidateWhereClause(whereClause);
        }

        var app = GetDatabase(filePath);
        dynamic? currentDb = null;
        dynamic? recordset = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            // Build SQL query
            var sql = $"SELECT * FROM [{tableName}]";
            if (!string.IsNullOrWhiteSpace(whereClause))
            {
                sql += $" WHERE {whereClause}";
            }

            _logger.LogDebug("Executing SQL: {Sql}", sql);

            try
            {
                recordset = currentDb.OpenRecordset(sql, 2); // dbOpenSnapshot = 2 (read-only)
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("not find") || ex.Message.Contains("exist"))
                {
                    throw new TableNotFoundException(tableName, filePath);
                }
                throw new QueryExecutionException($"Failed to query table: {ex.Message}", ex, filePath);
            }

            var result = RecordsetToTableData(recordset, limit, offset);

            if (result.ReturnedRows >= 500)
            {
                _logger.LogWarning("Large result set returned: {Count} rows from table {Table}",
                    (int)result.ReturnedRows, tableName);
            }

            _logger.LogInformation("Retrieved {Count} rows from table {Table}",
                (int)result.ReturnedRows, tableName);

            return result;
        }
        catch (TableNotFoundException)
        {
            throw;
        }
        catch (QueryExecutionException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting data from table {Table} in {Path}", tableName, filePath);
            throw new VbaOperationException($"Failed to get table data: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(recordset);
            ReleaseComObject(currentDb);
        }
    }

    #endregion

    #region Query Operations

    /// <summary>
    /// List all queries in the database
    /// </summary>
    public List<QueryInfo> ListQueries(string filePath)
    {
        var app = GetDatabase(filePath);
        var queries = new List<QueryInfo>();
        dynamic? currentDb = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            foreach (var queryDef in currentDb.QueryDefs)
            {
                try
                {
                    var name = queryDef.Name;

                    // Skip system queries
                    if (name.StartsWith("~") || name.StartsWith("MSys"))
                        continue;

                    var queryType = MapQueryType(queryDef.Type);
                    var paramCount = 0;

                    try
                    {
                        paramCount = queryDef.Parameters.Count;
                    }
                    catch
                    {
                        // Ignore
                    }

                    var queryInfo = new QueryInfo
                    {
                        Name = name,
                        QueryType = queryType,
                        DateCreated = TryGetDateTime(queryDef.DateCreated),
                        DateModified = TryGetDateTime(queryDef.LastUpdated),
                        ParameterCount = paramCount
                    };

                    queries.Add(queryInfo);
                    _logger.LogDebug("Found query: {Name} ({Type})", (string)name, (string)queryType);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read query info");
                }
            }

            _logger.LogInformation("Listed {Count} queries in {Path}", queries.Count, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing queries in {Path}", filePath);
            throw new VbaOperationException($"Failed to list queries: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(currentDb);
        }

        return queries;
    }

    /// <summary>
    /// Get SQL text of a query
    /// </summary>
    public string GetQuerySql(string filePath, string queryName)
    {
        var app = GetDatabase(filePath);
        dynamic? currentDb = null;
        dynamic? queryDef = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            try
            {
                queryDef = currentDb.QueryDefs[queryName];
            }
            catch
            {
                throw new QueryNotFoundException(queryName, filePath);
            }

            var sql = queryDef.SQL;
            _logger.LogInformation("Retrieved SQL for query {Query}", queryName);
            return sql;
        }
        catch (QueryNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting SQL for query {Query} in {Path}", queryName, filePath);
            throw new VbaOperationException($"Failed to get query SQL: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(queryDef);
            ReleaseComObject(currentDb);
        }
    }

    /// <summary>
    /// Execute a query
    /// </summary>
    public object ExecuteQuery(string filePath, string queryName,
        Dictionary<string, object>? parameters = null, int limit = 100, int offset = 0)
    {
        if (limit <= 0 || limit > 1000)
        {
            throw new ArgumentException("Limit must be between 1 and 1000", nameof(limit));
        }

        if (offset < 0)
        {
            throw new ArgumentException("Offset must be non-negative", nameof(offset));
        }

        var app = GetDatabase(filePath);
        dynamic? currentDb = null;
        dynamic? queryDef = null;
        dynamic? recordset = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            try
            {
                queryDef = currentDb.QueryDefs[queryName];
            }
            catch
            {
                throw new QueryNotFoundException(queryName, filePath);
            }

            // Set parameters if provided
            if (parameters != null && parameters.Count > 0)
            {
                foreach (dynamic param in queryDef.Parameters)
                {
                    try
                    {
                        string paramName = param.Name;
                        if (parameters.ContainsKey(paramName))
                        {
                            param.Value = parameters[paramName];
                            _logger.LogDebug("Set parameter {ParamName} = {Value}", paramName, parameters[paramName]);
                        }
                        else
                        {
                            _logger.LogWarning("Query parameter {ParamName} not provided", paramName);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(param);
                    }
                }
            }

            var queryType = queryDef.Type;
            const int dbQSelect = 0;

            if (queryType == dbQSelect)
            {
                // SELECT query - return data
                recordset = queryDef.OpenRecordset();
                var result = RecordsetToTableData(recordset, limit, offset);

                _logger.LogInformation("Executed SELECT query {Query}, returned {Count} rows",
                    queryName, (int)result.ReturnedRows);

                return result;
            }
            else
            {
                // Action query - execute and return affected rows
                queryDef.Execute();
                var recordsAffected = queryDef.RecordsAffected;

                var result = new QueryExecutionResult
                {
                    Success = true,
                    RecordsAffected = recordsAffected,
                    Message = $"Query executed successfully. {recordsAffected} record(s) affected."
                };

                _logger.LogInformation("Executed action query {Query}, {Count} records affected",
                    queryName, (int)recordsAffected);

                return result;
            }
        }
        catch (QueryNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error executing query {Query} in {Path}", queryName, filePath);
            throw new QueryExecutionException(ex.Message, ex, filePath);
        }
        finally
        {
            ReleaseComObject(recordset);
            ReleaseComObject(queryDef);
            ReleaseComObject(currentDb);
        }
    }

    /// <summary>
    /// Save (create or update) a query
    /// </summary>
    public void SaveQuery(string filePath, string queryName, string sql, bool replaceIfExists = false)
    {
        if (string.IsNullOrWhiteSpace(sql))
        {
            throw new ArgumentException("SQL cannot be empty", nameof(sql));
        }

        var app = GetDatabase(filePath);
        dynamic? currentDb = null;
        dynamic? queryDef = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            // Check if query exists
            var exists = false;
            try
            {
                var existing = currentDb.QueryDefs[queryName];
                exists = true;
                ReleaseComObject(existing);
            }
            catch
            {
                // Query doesn't exist
            }

            if (exists)
            {
                if (!replaceIfExists)
                {
                    throw new QueryAlreadyExistsException(queryName, filePath);
                }

                // Delete existing query
                currentDb.QueryDefs.Delete(queryName);
                _logger.LogDebug("Deleted existing query {Query}", queryName);
            }

            // Create new query
            queryDef = currentDb.CreateQueryDef(queryName, sql);

            _logger.LogInformation("Saved query {Query} in {Path}", queryName, filePath);
        }
        catch (QueryAlreadyExistsException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving query {Query} in {Path}", queryName, filePath);
            throw new VbaOperationException($"Failed to save query: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(queryDef);
            ReleaseComObject(currentDb);
        }
    }

    /// <summary>
    /// Delete a query
    /// </summary>
    public void DeleteQuery(string filePath, string queryName)
    {
        var app = GetDatabase(filePath);
        dynamic? currentDb = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            // Check if query exists
            try
            {
                var existing = currentDb.QueryDefs[queryName];
                ReleaseComObject(existing);
            }
            catch
            {
                throw new QueryNotFoundException(queryName, filePath);
            }

            // Delete query
            currentDb.QueryDefs.Delete(queryName);

            _logger.LogInformation("Deleted query {Query} from {Path}", queryName, filePath);
        }
        catch (QueryNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting query {Query} from {Path}", queryName, filePath);
            throw new VbaOperationException($"Failed to delete query: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(currentDb);
        }
    }

    #endregion

    #region Relationship Operations

    /// <summary>
    /// List all relationships in the database
    /// </summary>
    public List<RelationshipInfo> ListRelationships(string filePath)
    {
        var app = GetDatabase(filePath);
        var relationships = new List<RelationshipInfo>();
        dynamic? currentDb = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            foreach (dynamic relation in currentDb.Relations)
            {
                try
                {
                    dynamic name = relation.Name;
                    dynamic table = relation.Table;
                    dynamic foreignTable = relation.ForeignTable;
                    dynamic attributes = relation.Attributes;

                    // Get parent and child fields
                    var parentFields = new List<string>();
                    var childFields = new List<string>();

                    foreach (dynamic field in relation.Fields)
                    {
                        parentFields.Add((string)field.Name);
                        childFields.Add((string)field.ForeignName);
                        ReleaseComObject(field);
                    }

                    // Parse attributes flags
                    int attr = attributes;
                    bool enforceRI = (attr & 0x00000001) != 0;  // dbRelationUnique
                    bool cascadeUpdates = (attr & 0x00000100) != 0;  // dbRelationUpdateCascade
                    bool cascadeDeletes = (attr & 0x00001000) != 0;  // dbRelationDeleteCascade

                    // Determine relationship type
                    string relationType = (attr & 0x00000001) != 0 ? "One-to-One" : "One-to-Many";

                    var relationshipInfo = new RelationshipInfo
                    {
                        Name = (string)name,
                        ParentTable = (string)table,
                        ChildTable = (string)foreignTable,
                        ParentFields = parentFields,
                        ChildFields = childFields,
                        EnforceReferentialIntegrity = enforceRI,
                        CascadeUpdates = cascadeUpdates,
                        CascadeDeletes = cascadeDeletes,
                        RelationType = relationType
                    };

                    relationships.Add(relationshipInfo);
                    _logger.LogDebug("Found relationship: {Name} ({Parent} -> {Child})",
                        (string)name, (string)table, (string)foreignTable);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read relationship info");
                }
                finally
                {
                    ReleaseComObject(relation);
                }
            }

            _logger.LogInformation("Listed {Count} relationships in {Path}", relationships.Count, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing relationships in {Path}", filePath);
            throw new VbaOperationException($"Failed to list relationships: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(currentDb);
        }

        return relationships;
    }

    #endregion

    #region Index Operations

    /// <summary>
    /// Get all indexes for a table
    /// </summary>
    public List<IndexInfo> GetTableIndexes(string filePath, string tableName)
    {
        var app = GetDatabase(filePath);
        var indexes = new List<IndexInfo>();
        dynamic? currentDb = null;
        dynamic? tableDef = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            try
            {
                tableDef = currentDb.TableDefs[tableName];
            }
            catch
            {
                throw new TableNotFoundException(tableName, filePath);
            }

            foreach (dynamic index in tableDef.Indexes)
            {
                try
                {
                    dynamic name = index.Name;
                    dynamic primary = index.Primary;
                    dynamic unique = index.Unique;
                    dynamic foreign = index.Foreign;
                    dynamic ignoreNulls = index.IgnoreNulls;
                    dynamic required = index.Required;

                    // Get fields in the index
                    var fields = new List<string>();
                    foreach (dynamic field in index.Fields)
                    {
                        fields.Add((string)field.Name);
                        ReleaseComObject(field);
                    }

                    var indexInfo = new IndexInfo
                    {
                        Name = (string)name,
                        Fields = fields,
                        IsPrimary = primary,
                        IsUnique = unique,
                        IsForeign = foreign,
                        IgnoreNulls = ignoreNulls,
                        IsRequired = required
                    };

                    indexes.Add(indexInfo);
                    _logger.LogDebug("Found index: {Name} on fields [{Fields}]",
                        (string)name, string.Join(", ", fields));
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read index info");
                }
                finally
                {
                    ReleaseComObject(index);
                }
            }

            _logger.LogInformation("Retrieved {Count} indexes for table {Table}",
                indexes.Count, tableName);
        }
        catch (TableNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting indexes for table {Table} in {Path}",
                tableName, filePath);
            throw new VbaOperationException($"Failed to get table indexes: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(tableDef);
            ReleaseComObject(currentDb);
        }

        return indexes;
    }

    #endregion

    #region Database Information

    /// <summary>
    /// Get summary information about the database
    /// </summary>
    public DatabaseInfo GetDatabaseInfo(string filePath)
    {
        var app = GetDatabase(filePath);
        dynamic? currentDb = null;

        try
        {
            currentDb = app!.CurrentDb();
            if (currentDb == null)
            {
                throw new VbaOperationException("Failed to access database");
            }

            // Get file info
            var fileInfo = new FileInfo(filePath);

            // Count objects
            int tableCount = 0;
            foreach (dynamic tableDef in currentDb.TableDefs)
            {
                string name = tableDef.Name;
                int attrs = tableDef.Attributes;
                // Skip system and temporary tables
                if (!name.StartsWith("MSys") && !name.StartsWith("~") && (attrs & 0x80000000) == 0)
                {
                    tableCount++;
                }
                ReleaseComObject(tableDef);
            }

            int queryCount = 0;
            foreach (dynamic queryDef in currentDb.QueryDefs)
            {
                string name = queryDef.Name;
                if (!name.StartsWith("~"))
                {
                    queryCount++;
                }
                ReleaseComObject(queryDef);
            }

            int relationshipCount = 0;
            foreach (dynamic relation in currentDb.Relations)
            {
                relationshipCount++;
                ReleaseComObject(relation);
            }

            // Count forms and reports (via Application.AllForms/AllReports)
            int formCount = 0;
            int reportCount = 0;
            try
            {
                foreach (dynamic formObj in app.CurrentProject.AllForms)
                {
                    formCount++;
                    ReleaseComObject(formObj);
                }
            }
            catch
            {
                // Forms collection might not be accessible
            }

            try
            {
                foreach (dynamic reportObj in app.CurrentProject.AllReports)
                {
                    reportCount++;
                    ReleaseComObject(reportObj);
                }
            }
            catch
            {
                // Reports collection might not be accessible
            }

            // Get database version
            string version = "Unknown";
            try
            {
                version = currentDb.Version;
            }
            catch
            {
                // Version might not be accessible
            }

            var dbInfo = new DatabaseInfo
            {
                FilePath = filePath,
                FileSizeBytes = fileInfo.Length,
                FileSizeFormatted = FormatFileSize(fileInfo.Length),
                Version = version,
                TableCount = tableCount,
                QueryCount = queryCount,
                FormCount = formCount,
                ReportCount = reportCount,
                RelationshipCount = relationshipCount,
                DateCreated = fileInfo.CreationTime,
                DateModified = fileInfo.LastWriteTime,
                IsPasswordProtected = false  // Cannot easily detect password protection
            };

            _logger.LogInformation("Retrieved database info for {Path}: {Tables} tables, {Queries} queries",
                filePath, tableCount, queryCount);

            return dbInfo;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting database info for {Path}", filePath);
            throw new VbaOperationException($"Failed to get database info: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(currentDb);
        }
    }

    /// <summary>
    /// List all forms in the database
    /// </summary>
    public List<DatabaseObjectInfo> ListForms(string filePath)
    {
        var app = GetDatabase(filePath);
        var forms = new List<DatabaseObjectInfo>();

        try
        {
            foreach (dynamic formObj in app!.CurrentProject.AllForms)
            {
                try
                {
                    var objectInfo = new DatabaseObjectInfo
                    {
                        Name = (string)formObj.Name,
                        Type = "Form",
                        DateCreated = TryGetDateTime(formObj.DateCreated),
                        DateModified = TryGetDateTime(formObj.DateModified),
                        IsLoaded = formObj.IsLoaded
                    };

                    forms.Add(objectInfo);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read form info");
                }
                finally
                {
                    ReleaseComObject(formObj);
                }
            }

            _logger.LogInformation("Listed {Count} forms in {Path}", forms.Count, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing forms in {Path}", filePath);
            throw new VbaOperationException($"Failed to list forms: {ex.Message}", ex);
        }

        return forms;
    }

    /// <summary>
    /// List all reports in the database
    /// </summary>
    public List<DatabaseObjectInfo> ListReports(string filePath)
    {
        var app = GetDatabase(filePath);
        var reports = new List<DatabaseObjectInfo>();

        try
        {
            foreach (dynamic reportObj in app!.CurrentProject.AllReports)
            {
                try
                {
                    var objectInfo = new DatabaseObjectInfo
                    {
                        Name = (string)reportObj.Name,
                        Type = "Report",
                        DateCreated = TryGetDateTime(reportObj.DateCreated),
                        DateModified = TryGetDateTime(reportObj.DateModified),
                        IsLoaded = reportObj.IsLoaded
                    };

                    reports.Add(objectInfo);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read report info");
                }
                finally
                {
                    ReleaseComObject(reportObj);
                }
            }

            _logger.LogInformation("Listed {Count} reports in {Path}", reports.Count, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing reports in {Path}", filePath);
            throw new VbaOperationException($"Failed to list reports: {ex.Message}", ex);
        }

        return reports;
    }

    #endregion

    #region Export Operations

    /// <summary>
    /// Export table data to a CSV file
    /// </summary>
    public void ExportTableToCsv(string filePath, string tableName, string outputPath,
        string? whereClause = null, int limit = 0)
    {
        _logger.LogInformation("Exporting table {Table} to CSV: {Output}", tableName, outputPath);

        // Get table data
        var data = GetTableData(filePath, tableName, whereClause, limit, 0);

        // Write to file
        using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
        writer.Write(FormatAsCsv(data));

        _logger.LogInformation("Exported {Count} rows to {Output}", data.ReturnedRows, outputPath);
    }

    /// <summary>
    /// Export query results to a CSV file
    /// </summary>
    public void ExportQueryToCsv(string filePath, string queryName, string outputPath,
        Dictionary<string, object>? parameters = null, int limit = 0)
    {
        _logger.LogInformation("Exporting query {Query} to CSV: {Output}", queryName, outputPath);

        // Execute query
        var result = ExecuteQuery(filePath, queryName, parameters, limit, 0);

        if (result is TableDataResult data)
        {
            // Write to file
            using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
            writer.Write(FormatAsCsv(data));

            _logger.LogInformation("Exported {Count} rows to {Output}", data.ReturnedRows, outputPath);
        }
        else
        {
            throw new VbaOperationException("Query did not return data (it may be an action query)");
        }
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Map DAO data type to friendly string
    /// </summary>
    private string MapDataType(int dataType)
    {
        return dataType switch
        {
            1 => "Yes/No",           // dbBoolean
            2 => "Byte",             // dbByte
            3 => "Number (Integer)", // dbInteger
            4 => "Number (Long)",    // dbLong
            5 => "Currency",         // dbCurrency
            6 => "Number (Single)",  // dbSingle
            7 => "Number (Double)",  // dbDouble
            8 => "Date/Time",        // dbDate
            10 => "Text",            // dbText
            11 => "OLE Object",      // dbLongBinary
            12 => "Memo",            // dbMemo
            15 => "Number (ReplicationID)", // dbGUID
            16 => "Number (Big Integer)",   // dbBigInt
            17 => "VarBinary",       // dbVarBinary
            18 => "Char",            // dbChar
            19 => "Number (Numeric/Decimal)", // dbNumeric
            20 => "Number (Decimal)", // dbDecimal
            21 => "Number (Float)",   // dbFloat
            22 => "Time",             // dbTime
            23 => "TimeStamp",        // dbTimeStamp
            _ => $"Unknown ({dataType})"
        };
    }

    /// <summary>
    /// Map query type to friendly string
    /// </summary>
    private string MapQueryType(int queryType)
    {
        return queryType switch
        {
            0 => "Select",        // dbQSelect
            48 => "Crosstab",     // dbQCrosstab
            80 => "Delete",       // dbQDelete
            112 => "Update",      // dbQUpdate
            64 => "Append",       // dbQAppend
            96 => "Make Table",   // dbQMakeTable
            128 => "DDL",         // dbQDDL
            240 => "SQL Passthrough", // dbQSQLPassthrough
            256 => "Union",       // dbQSetOperation
            _ => $"Unknown ({queryType})"
        };
    }

    /// <summary>
    /// Get table type based on attributes
    /// </summary>
    private string GetTableType(int attributes)
    {
        const int dbAttachedTable = 0x40000000;
        const int dbSystemObject = unchecked((int)0x80000002);

        if ((attributes & dbSystemObject) != 0)
            return "System";
        if ((attributes & dbAttachedTable) != 0)
            return "LinkedTable";

        return "Table";
    }

    /// <summary>
    /// Validate WHERE clause for SQL injection
    /// </summary>
    private void ValidateWhereClause(string whereClause)
    {
        var lowerClause = whereClause.ToLowerInvariant();

        var prohibitedKeywords = new[]
        {
            "drop ", "delete ", "insert ", "create ", "alter ",
            "exec ", "execute ", "truncate ", "--", "/*", "*/", "xp_"
        };

        foreach (var keyword in prohibitedKeywords)
        {
            if (lowerClause.Contains(keyword))
            {
                throw new InvalidSqlException(
                    $"WHERE clause contains prohibited keyword '{keyword.Trim()}'. " +
                    "Use parameterized queries for complex operations.");
            }
        }

        // Check for semicolon (statement separator)
        if (whereClause.Contains(';'))
        {
            throw new InvalidSqlException(
                "WHERE clause contains statement separator ';'. This is not allowed.");
        }
    }

    /// <summary>
    /// Convert recordset to TableDataResult
    /// </summary>
    private TableDataResult RecordsetToTableData(dynamic recordset, int limit, int offset)
    {
        var columnNames = new List<string>();
        var rows = new List<Dictionary<string, object?>>();

        // Get column names
        foreach (var field in recordset.Fields)
        {
            columnNames.Add(field.Name);
        }

        // Count total rows
        var totalRows = 0;
        if (!recordset.EOF)
        {
            recordset.MoveLast();
            totalRows = recordset.RecordCount;
            recordset.MoveFirst();
        }

        // Skip offset rows
        if (offset > 0 && !recordset.EOF)
        {
            recordset.Move(offset);
        }

        // Read data rows
        var count = 0;
        while (!recordset.EOF && count < limit)
        {
            var row = new Dictionary<string, object?>();
            foreach (var field in recordset.Fields)
            {
                var value = field.Value;
                row[field.Name] = ConvertDbValue(value);
            }
            rows.Add(row);
            count++;
            recordset.MoveNext();
        }

        var hasMore = !recordset.EOF;

        return new TableDataResult
        {
            ColumnNames = columnNames,
            Rows = rows,
            TotalRows = totalRows,
            ReturnedRows = rows.Count,
            HasMore = hasMore
        };
    }

    /// <summary>
    /// Convert database value to .NET type
    /// </summary>
    private object? ConvertDbValue(object? value)
    {
        if (value == null || value is DBNull)
            return null;

        // Convert DateTime to ISO 8601 string for JSON serialization
        if (value is DateTime dt)
            return dt.ToString("yyyy-MM-ddTHH:mm:ss");

        return value;
    }

    /// <summary>
    /// Format table data as CSV
    /// </summary>
    public string FormatAsCsv(TableDataResult data)
    {
        var csv = new System.Text.StringBuilder();

        // Header row
        csv.AppendLine(string.Join(",", data.ColumnNames.Select(EscapeCsvField)));

        // Data rows
        foreach (var row in data.Rows)
        {
            var values = data.ColumnNames.Select(col =>
            {
                var value = row.ContainsKey(col) ? row[col] : null;
                return EscapeCsvField(value?.ToString() ?? "");
            });
            csv.AppendLine(string.Join(",", values));
        }

        return csv.ToString();
    }

    /// <summary>
    /// Escape a field value for CSV format
    /// </summary>
    private string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field))
            return "";

        // Escape if contains comma, quote, or newline
        if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }

        return field;
    }

    /// <summary>
    /// Try to get DateTime from COM object
    /// </summary>
    private DateTime? TryGetDateTime(object? value)
    {
        if (value == null || value is DBNull)
            return null;

        try
        {
            return Convert.ToDateTime(value);
        }
        catch
        {
            return null;
        }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to release COM object");
            }
        }
    }

    /// <summary>
    /// Format file size in bytes to human-readable string
    /// </summary>
    private string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB", "TB" };
        double len = bytes;
        int order = 0;

        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }

        return $"{len:0.##} {sizes[order]}";
    }

    #endregion
}
