using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
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
    private readonly BackupService _backupService;

    public ExcelVbaTools(ExcelComService excelService, BackupService backupService)
    {
        _excelService = excelService;
        _backupService = backupService;
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

    [McpServerTool(Name = "list_vba_modules")]
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

    [McpServerTool(Name = "read_vba_module")]
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

    [McpServerTool(Name = "write_vba_module")]
    [Description("Write VBA code to a module, replacing its entire content. Creates a backup automatically. The workbook must be open in Excel.")]
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
            // Create backup first
            var existingCode = string.Empty;
            try
            {
                existingCode = _excelService.ReadModule(filePath, moduleName);
            }
            catch
            {
                // Module might be empty or new
            }

            string? backupPath = null;
            if (!string.IsNullOrEmpty(existingCode))
            {
                backupPath = _backupService.BackupModule(filePath, moduleName, existingCode);
            }

            // Write new code
            _excelService.WriteModule(filePath, moduleName, code);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                linesWritten = code.Split('\n').Length,
                backupCreated = backupPath != null,
                backupPath = backupPath
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

    [McpServerTool(Name = "create_vba_module")]
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

    [McpServerTool(Name = "delete_vba_module")]
    [Description("Delete a VBA module from an Excel workbook. Creates a backup automatically. Cannot delete document modules (ThisWorkbook, Sheet modules).")]
    public string DeleteVbaModule(
        [Description("Full file path to the Excel workbook")] 
        string filePath,
        [Description("Name of the module to delete")] 
        string moduleName)
    {
        try
        {
            // Backup before deletion
            var existingCode = _excelService.ReadModule(filePath, moduleName);
            var backupPath = _backupService.BackupModule(filePath, moduleName, existingCode);

            _excelService.DeleteModule(filePath, moduleName);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                deleted = true,
                backupPath = backupPath
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

    [McpServerTool(Name = "export_vba_module")]
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

    [McpServerTool(Name = "list_vba_backups")]
    [Description("List all VBA code backups created by this tool")]
    public string ListVbaBackups(
        [Description("Optional: Filter backups by source file path")] 
        string? filePath = null)
    {
        var backups = _backupService.ListBackups(filePath);

        var result = new
        {
            count = backups.Count,
            backups = backups.Select(b => new
            {
                fileName = b.FileName,
                createdAt = b.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"),
                sizeBytes = b.SizeBytes,
                fullPath = b.FullPath
            })
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}
