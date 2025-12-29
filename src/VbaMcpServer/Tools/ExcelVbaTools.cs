using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using VbaMcpServer.Exceptions;
using VbaMcpServer.Models;
using VbaMcpServer.Services;

namespace VbaMcpServer.Tools;

/// <summary>
/// MCP Tools for Excel VBA operations
/// </summary>
[McpServerToolType]
public class ExcelVbaTools
{
    private readonly ExcelComService _excelService;

    public ExcelVbaTools(ExcelComService excelService)
    {
        _excelService = excelService;
    }

    [McpServerTool(Name = "list_open_excel_files")]
    [Description("List all currently open Excel workbooks that contain VBA projects (.xlsm, .xlsb, .xls)")]
    public string ListOpenExcelFiles()
    {
        var workbooks = _excelService.ListOpenWorkbooks();
        
        if (workbooks.Count == 0)
        {
            return "No Excel workbooks are currently open, or Excel is not running.";
        }

        var result = new
        {
            count = workbooks.Count,
            workbooks = workbooks
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool(Name = "list_excel_vba_modules")]
    [Description("List all VBA modules in an Excel workbook. The workbook must be open in Excel.")]
    public string ListVbaModules(
        [Description("Full file path to the Excel workbook (e.g., C:\\Projects\\MyWorkbook.xlsm)")] 
        string filePath)
    {
        try
        {
            var modules = _excelService.ListModules(filePath);
            
            var result = new
            {
                file = filePath,
                moduleCount = modules.Count,
                modules = modules
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}. Please open the file in Excel first.";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "read_excel_vba_module")]
    [Description("Read the complete VBA code from a module. The workbook must be open in Excel.")]
    public string ReadVbaModule(
        [Description("Full file path to the Excel workbook")] 
        string filePath,
        [Description("Name of the VBA module to read (e.g., Module1, Sheet1, ThisWorkbook)")] 
        string moduleName)
    {
        try
        {
            var code = _excelService.ReadModule(filePath, moduleName);
            
            if (string.IsNullOrEmpty(code))
            {
                return $"Module '{moduleName}' exists but contains no code.";
            }

            var result = new
            {
                file = filePath,
                module = moduleName,
                lineCount = code.Split('\n').Length,
                code = code
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "write_excel_vba_module")]
    [Description("Write VBA code to a module, replacing its entire content. IMPORTANT: This operation is irreversible. Make sure to backup your file before using this tool. The workbook must be open in Excel.")]
    public string WriteVbaModule(
        [Description("Full file path to the Excel workbook")]
        string filePath,
        [Description("Name of the VBA module to write to")]
        string moduleName,
        [Description("The complete VBA code to write to the module")]
        string code)
    {
        try
        {
            // Write new code
            _excelService.WriteModule(filePath, moduleName, code);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                linesWritten = code.Split('\n').Length
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "create_excel_vba_module")]
    [Description("Create a new VBA module in an Excel workbook. The workbook must be open in Excel.")]
    public string CreateVbaModule(
        [Description("Full file path to the Excel workbook")] 
        string filePath,
        [Description("Name for the new module")] 
        string moduleName,
        [Description("Type of module: 'standard' (default), 'class', or 'userform'")] 
        string moduleType = "standard")
    {
        try
        {
            var type = moduleType.ToLowerInvariant() switch
            {
                "standard" or "std" => VbaModuleType.StdModule,
                "class" => VbaModuleType.ClassModule,
                "userform" or "form" => VbaModuleType.MSForm,
                _ => VbaModuleType.StdModule
            };

            _excelService.CreateModule(filePath, moduleName, type);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                type = type.ToString()
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "delete_excel_vba_module")]
    [Description("Delete a VBA module from an Excel workbook. IMPORTANT: This operation is irreversible. Make sure to backup your file before using this tool. Cannot delete document modules (ThisWorkbook, Sheet modules).")]
    public string DeleteVbaModule(
        [Description("Full file path to the Excel workbook")]
        string filePath,
        [Description("Name of the module to delete")]
        string moduleName)
    {
        try
        {
            _excelService.DeleteModule(filePath, moduleName);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                deleted = true
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (InvalidOperationException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "export_excel_vba_module")]
    [Description("Export a VBA module to a file (.bas, .cls, or .frm)")]
    public string ExportVbaModule(
        [Description("Full file path to the Excel workbook")] 
        string filePath,
        [Description("Name of the module to export")] 
        string moduleName,
        [Description("Output file path (e.g., C:\\Exports\\Module1.bas)")] 
        string outputPath)
    {
        try
        {
            _excelService.ExportModule(filePath, moduleName, outputPath);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                exportedTo = outputPath
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "list_excel_vba_procedures")]
    [Description("List all procedures in a VBA module with detailed metadata including name, type, line numbers, and access modifiers")]
    public string ListVbaProcedures(
        [Description("Full file path to the Excel workbook (e.g., C:\\MyWorkbook.xlsm)")] string filePath,
        [Description("Name of the VBA module to list procedures from")] string moduleName)
    {
        try
        {
            var procedures = _excelService.ListProcedures(filePath, moduleName);

            var result = new
            {
                file = filePath,
                module = moduleName,
                procedureCount = procedures.Count,
                procedures = procedures.Select(p => new
                {
                    name = p.Name,
                    type = p.Type,
                    startLine = p.StartLine,
                    lineCount = p.LineCount,
                    accessModifier = p.AccessModifier
                })
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Excel's Trust Center settings.";
        }
        catch (ModuleNotFoundException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error listing procedures: {ex.Message}";
        }
    }

    [McpServerTool(Name = "read_excel_vba_procedure")]
    [Description("Read the code of a specific procedure from a VBA module")]
    public string ReadVbaProcedure(
        [Description("Full file path to the Excel workbook (e.g., C:\\MyWorkbook.xlsm)")] string filePath,
        [Description("Name of the VBA module containing the procedure")] string moduleName,
        [Description("Name of the procedure to read (Sub, Function, or Property)")] string procedureName)
    {
        try
        {
            var code = _excelService.ReadProcedure(filePath, moduleName, procedureName);

            var result = new
            {
                file = filePath,
                module = moduleName,
                procedure = procedureName,
                lineCount = code.Split('\n').Length,
                code = code
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Excel's Trust Center settings.";
        }
        catch (ModuleNotFoundException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error reading procedure: {ex.Message}";
        }
    }

    [McpServerTool(Name = "write_excel_vba_procedure")]
    [Description("Write or replace a specific procedure in a VBA module. IMPORTANT: This operation is irreversible. The procedure will be replaced with the new code.")]
    public string WriteVbaProcedure(
        [Description("Full file path to the Excel workbook (e.g., C:\\MyWorkbook.xlsm)")] string filePath,
        [Description("Name of the VBA module containing the procedure")] string moduleName,
        [Description("Name of the procedure to write/replace (Sub, Function, or Property)")] string procedureName,
        [Description("The complete VBA code for the procedure, including the procedure declaration (Sub/Function/Property) and End statement")] string code)
    {
        try
        {
            _excelService.WriteProcedure(filePath, moduleName, procedureName, code);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                procedure = procedureName,
                linesWritten = code.Split('\n').Length,
                warning = "This operation is irreversible. The procedure has been replaced."
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Workbook not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Excel's Trust Center settings.";
        }
        catch (ModuleNotFoundException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error writing procedure: {ex.Message}";
        }
    }
}
