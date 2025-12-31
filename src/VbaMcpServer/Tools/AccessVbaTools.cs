using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using VbaMcpServer.Exceptions;
using VbaMcpServer.Helpers;
using VbaMcpServer.Models;
using VbaMcpServer.Services;

namespace VbaMcpServer.Tools;

/// <summary>
/// MCP Tools for Access VBA operations
/// </summary>
[McpServerToolType]
public class AccessVbaTools
{
    private readonly AccessComService _accessService;

    public AccessVbaTools(AccessComService accessService)
    {
        _accessService = accessService;
    }

    [McpServerTool(Name = "list_open_access_files")]
    [Description("List all currently open Access databases that contain VBA projects (.accdb, .mdb)")]
    public string ListOpenAccessFiles()
    {
        var databases = _accessService.ListOpenDatabases();

        if (databases.Count == 0)
        {
            return "No Access databases are currently open, or Access is not running.";
        }

        var result = new
        {
            count = databases.Count,
            databases = databases
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool(Name = "list_access_vba_modules")]
    [Description("List all VBA modules in an Access database. The database must be open in Access.")]
    public string ListAccessVbaModules(
        [Description("Full file path to the Access database (e.g., C:\\Projects\\MyDatabase.accdb)")]
        string filePath)
    {
        try
        {
            var modules = _accessService.ListModules(filePath);

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
            return $"Error: Database not found or not open: {filePath}. Please open the file in Access first.";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "read_access_vba_module")]
    [Description("Read the complete VBA code from a module in an Access database. The database must be open in Access.")]
    public string ReadAccessVbaModule(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the VBA module to read (e.g., Module1, Form_MainForm, Report_Report1)")]
        string moduleName)
    {
        try
        {
            var code = _accessService.ReadModule(filePath, moduleName);

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
            return $"Error: Database not found or not open: {filePath}";
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

    [McpServerTool(Name = "write_access_vba_module")]
    [Description("Write VBA code to a module in an Access database, replacing its entire content. IMPORTANT: This operation is irreversible. Make sure to backup your file before using this tool. The database must be open in Access.")]
    public string WriteAccessVbaModule(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the VBA module to write to")]
        string moduleName,
        [Description("The complete VBA code to write to the module")]
        string code)
    {
        try
        {
            // Write new code
            _accessService.WriteModule(filePath, moduleName, code);

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
            return $"Error: Database not found or not open: {filePath}";
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

    [McpServerTool(Name = "create_access_vba_module")]
    [Description("Create a new VBA module in an Access database. The database must be open in Access.")]
    public string CreateAccessVbaModule(
        [Description("Full file path to the Access database")]
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

            _accessService.CreateModule(filePath, moduleName, type);

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
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "delete_access_vba_module")]
    [Description("Delete a VBA module from an Access database. IMPORTANT: This operation is irreversible. Make sure to backup your file before using this tool. Cannot delete document modules (Forms, Reports with code-behind).")]
    public string DeleteAccessVbaModule(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the module to delete")]
        string moduleName)
    {
        try
        {
            _accessService.DeleteModule(filePath, moduleName);

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
            return $"Error: Database not found or not open: {filePath}";
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

    [McpServerTool(Name = "export_access_vba_module")]
    [Description("Export a VBA module from an Access database to a file (.bas, .cls, or .frm)")]
    public string ExportAccessVbaModule(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the module to export")]
        string moduleName,
        [Description("Output file path (e.g., C:\\Exports\\Module1.bas)")]
        string outputPath)
    {
        try
        {
            _accessService.ExportModule(filePath, moduleName, outputPath);

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
            return $"Error: Database not found or not open: {filePath}";
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

    [McpServerTool(Name = "list_access_vba_procedures")]
    [Description("List all procedures in an Access VBA module with detailed metadata including name, type, line numbers, and access modifiers")]
    public string ListAccessVbaProcedures(
        [Description("Full file path to the Access database (e.g., C:\\MyDatabase.accdb)")] string filePath,
        [Description("Name of the VBA module to list procedures from")] string moduleName)
    {
        try
        {
            var procedures = _accessService.ListProcedures(filePath, moduleName);

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
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Access's Trust Center settings.";
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

    [McpServerTool(Name = "read_access_vba_procedure")]
    [Description("Read the code of a specific procedure from an Access VBA module")]
    public string ReadAccessVbaProcedure(
        [Description("Full file path to the Access database (e.g., C:\\MyDatabase.accdb)")] string filePath,
        [Description("Name of the VBA module containing the procedure")] string moduleName,
        [Description("Name of the procedure to read (Sub, Function, or Property)")] string procedureName)
    {
        try
        {
            var code = _accessService.ReadProcedure(filePath, moduleName, procedureName);

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
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Access's Trust Center settings.";
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

    [McpServerTool(Name = "write_access_vba_procedure")]
    [Description("Write or add a procedure in an Access VBA module. If the procedure exists, it will be replaced. If it does not exist, it will be added to the end of the module. IMPORTANT: This operation is irreversible. Do NOT apply XML escaping to the code (use '&' not '&amp;', '<' not '&lt;').")]
    public string WriteAccessVbaProcedure(
        [Description("Full file path to the Access database (e.g., C:\\MyDatabase.accdb)")] string filePath,
        [Description("Name of the VBA module containing the procedure")] string moduleName,
        [Description("Name of the procedure to write/replace (Sub, Function, or Property)")] string procedureName,
        [Description("The complete VBA code for the procedure, including the procedure declaration (Sub/Function/Property) and End statement. IMPORTANT: Do NOT apply XML escaping (use '&' not '&amp;', '<' not '&lt;', '>' not '&gt;')")] string code)
    {
        try
        {
            var action = _accessService.WriteProcedure(filePath, moduleName, procedureName, code);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                procedure = procedureName,
                action = action,
                linesWritten = code.Split('\n').Length,
                message = action == "replaced"
                    ? "Procedure replaced successfully"
                    : "Procedure added successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Access's Trust Center settings.";
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

    #region Table Tools

    [McpServerTool(Name = "list_access_tables")]
    [Description("List all tables in an Access database")]
    public string ListAccessTables(
        [Description("Full file path to the Access database (e.g., C:\\Projects\\MyDatabase.accdb)")] string filePath,
        [Description("Include system tables (default: false)")] bool includeSystemTables = false,
        [Description("Output format: 'json' or 'csv' (default: 'json')")] string format = "json")
    {
        try
        {
            var tables = _accessService.ListTables(filePath, includeSystemTables);

            if (format.ToLowerInvariant() == "csv")
            {
                var csv = new System.Text.StringBuilder();
                csv.AppendLine("Name,Type,RecordCount,DateCreated,DateModified");
                foreach (var table in tables)
                {
                    csv.AppendLine($"{table.Name},{table.Type},{table.RecordCount},{table.DateCreated},{table.DateModified}");
                }
                return csv.ToString();
            }

            var result = new
            {
                file = filePath,
                tableCount = tables.Count,
                tables = tables
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}. Please open the file in Access first.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "get_access_table_structure")]
    [Description("Get the structure (field definitions) of an Access table")]
    public string GetAccessTableStructure(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Table name")] string tableName,
        [Description("Output format: 'json' or 'csv' (default: 'json')")] string format = "json")
    {
            try
        {
            var fields = _accessService.GetTableStructure(filePath, tableName);

            if (format.ToLowerInvariant() == "csv")
            {
                var csv = new System.Text.StringBuilder();
                csv.AppendLine("Name,DataType,Size,Required,AllowZeroLength,DefaultValue,ValidationRule,IsPrimaryKey,IsIndexed");
                foreach (var field in fields)
                {
                    csv.AppendLine($"{field.Name},{field.DataType},{field.Size},{field.Required},{field.AllowZeroLength},{field.DefaultValue},{field.ValidationRule},{field.IsPrimaryKey},{field.IsIndexed}");
                }
                return csv.ToString();
            }

            var result = new
            {
                file = filePath,
                table = tableName,
                fieldCount = fields.Count,
                fields = fields
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (TableNotFoundException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "get_access_table_data")]
    [Description("Get data from an Access table with optional WHERE clause and pagination")]
    public string GetAccessTableData(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Table name")] string tableName,
        [Description("WHERE clause without 'WHERE' keyword (optional)")] string? whereClause = null,
        [Description("Maximum rows to return (default: 100, max: 1000)")] int limit = 100,
        [Description("Number of rows to skip (default: 0)")] int offset = 0,
        [Description("Output format: 'json' or 'csv' (default: 'json')")] string format = "json")
    {
        try
        {
            var data = _accessService.GetTableData(filePath, tableName, whereClause, limit, offset);

            if (format.ToLowerInvariant() == "csv")
            {
                return _accessService.FormatAsCsv(data);
            }

            var result = new
            {
                file = filePath,
                table = tableName,
                whereClause = whereClause,
                totalRows = data.TotalRows,
                returnedRows = data.ReturnedRows,
                hasMore = data.HasMore,
                data = data
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (TableNotFoundException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (InvalidSqlException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    #endregion

    #region Query Tools

    [McpServerTool(Name = "list_access_queries")]
    [Description("List all saved queries in an Access database")]
    public string ListAccessQueries(
        [Description("Full file path to the Access database (e.g., C:\\Projects\\MyDatabase.accdb)")] string filePath,
        [Description("Output format: 'json' or 'csv' (default: 'json')")] string format = "json")
    {
        try
        {
            var queries = _accessService.ListQueries(filePath);

            if (format.ToLowerInvariant() == "csv")
            {
                var csv = new System.Text.StringBuilder();
                csv.AppendLine("Name,QueryType,ParameterCount,DateCreated,DateModified");
                foreach (var query in queries)
                {
                    csv.AppendLine($"{query.Name},{query.QueryType},{query.ParameterCount},{query.DateCreated},{query.DateModified}");
                }
                return csv.ToString();
            }

            var result = new
            {
                file = filePath,
                queryCount = queries.Count,
                queries = queries
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}. Please open the file in Access first.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "get_access_query_sql")]
    [Description("Get the SQL text of a saved Access query")]
    public string GetAccessQuerySql(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Query name")] string queryName)
    {
        try
        {
            var sql = _accessService.GetQuerySql(filePath, queryName);

            var result = new
            {
                file = filePath,
                query = queryName,
                sql = sql
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (QueryNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_queries to see available queries.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "execute_access_query")]
    [Description("Execute a saved Access query and return the results")]
    public string ExecuteAccessQuery(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Query name")] string queryName,
        [Description("Query parameters as JSON object (optional)")] string? parametersJson = null,
        [Description("Maximum rows to return for SELECT queries (default: 100)")] int limit = 100,
        [Description("Number of rows to skip (default: 0)")] int offset = 0,
        [Description("Output format: 'json' or 'csv' (default: 'json')")] string format = "json")
    {
        try
        {
            Dictionary<string, object>? parameters = null;
            if (!string.IsNullOrWhiteSpace(parametersJson))
            {
                parameters = JsonSerializer.Deserialize<Dictionary<string, object>>(parametersJson);
            }

            var queryResult = _accessService.ExecuteQuery(filePath, queryName, parameters, limit, offset);

            if (queryResult is TableDataResult tableData)
            {
                if (format.ToLowerInvariant() == "csv")
                {
                    return _accessService.FormatAsCsv(tableData);
                }

                var result = new
                {
                    file = filePath,
                    query = queryName,
                    resultType = "data",
                    totalRows = tableData.TotalRows,
                    returnedRows = tableData.ReturnedRows,
                    hasMore = tableData.HasMore,
                    data = tableData
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else if (queryResult is QueryExecutionResult execResult)
            {
                var result = new
                {
                    file = filePath,
                    query = queryName,
                    resultType = "execution",
                    success = execResult.Success,
                    recordsAffected = execResult.RecordsAffected,
                    message = execResult.Message
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }

            return "Error: Unknown result type";
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (QueryNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_queries to see available queries.";
        }
        catch (QueryExecutionException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "save_access_query")]
    [Description("Save (create or update) a query in an Access database")]
    public string SaveAccessQuery(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Query name")] string queryName,
        [Description("SQL statement")] string sql,
        [Description("Replace if exists (default: false)")] bool replaceIfExists = false)
    {
        try
        {
            _accessService.SaveQuery(filePath, queryName, sql, replaceIfExists);

            var result = new
            {
                success = true,
                file = filePath,
                query = queryName,
                message = replaceIfExists ? "Query replaced successfully" : "Query created successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (QueryAlreadyExistsException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "delete_access_query")]
    [Description("Delete a saved query from an Access database")]
    public string DeleteAccessQuery(
        [Description("Full file path to the Access database")] string filePath,
        [Description("Query name")] string queryName)
    {
        try
        {
            _accessService.DeleteQuery(filePath, queryName);

            var result = new
            {
                success = true,
                file = filePath,
                query = queryName,
                message = "Query deleted successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (QueryNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_queries to see available queries.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    #endregion

    #region Relationship and Index Tools

    [McpServerTool(Name = "list_access_relationships")]
    [Description("List all relationships between tables in an Access database")]
    public string ListAccessRelationships(
        [Description("Full file path to the Access database")]
        string filePath)
    {
        try
        {
            var relationships = _accessService.ListRelationships(filePath);

            if (relationships.Count == 0)
            {
                return "No relationships found in the database.";
            }

            var result = new
            {
                count = relationships.Count,
                relationships = relationships
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "get_access_table_indexes")]
    [Description("Get all indexes for a specific table in an Access database")]
    public string GetAccessTableIndexes(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the table")]
        string tableName)
    {
        try
        {
            var indexes = _accessService.GetTableIndexes(filePath, tableName);

            if (indexes.Count == 0)
            {
                return $"No indexes found for table '{tableName}'.";
            }

            var result = new
            {
                table = tableName,
                count = indexes.Count,
                indexes = indexes
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (TableNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_tables to see available tables.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    #endregion

    #region Database Information Tools

    [McpServerTool(Name = "get_access_database_info")]
    [Description("Get summary information about an Access database (file size, table count, query count, etc.)")]
    public string GetAccessDatabaseInfo(
        [Description("Full file path to the Access database")]
        string filePath)
    {
        try
        {
            var info = _accessService.GetDatabaseInfo(filePath);
            return JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "list_access_forms")]
    [Description("List all forms in an Access database")]
    public string ListAccessForms(
        [Description("Full file path to the Access database")]
        string filePath)
    {
        try
        {
            var forms = _accessService.ListForms(filePath);

            if (forms.Count == 0)
            {
                return "No forms found in the database.";
            }

            var result = new
            {
                count = forms.Count,
                forms = forms
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "list_access_reports")]
    [Description("List all reports in an Access database")]
    public string ListAccessReports(
        [Description("Full file path to the Access database")]
        string filePath)
    {
        try
        {
            var reports = _accessService.ListReports(filePath);

            if (reports.Count == 0)
            {
                return "No reports found in the database.";
            }

            var result = new
            {
                count = reports.Count,
                reports = reports
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    #endregion

    #region Export Tools

    [McpServerTool(Name = "export_access_table_to_csv")]
    [Description("Export Access table data to a CSV file")]
    public string ExportAccessTableToCsv(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the table to export")]
        string tableName,
        [Description("Full file path for the output CSV file")]
        string outputPath,
        [Description("Optional WHERE clause to filter data (without 'WHERE' keyword)")]
        string? whereClause = null,
        [Description("Maximum number of rows to export (0 = unlimited, default: 0)")]
        int limit = 0)
    {
        try
        {
            _accessService.ExportTableToCsv(filePath, tableName, outputPath, whereClause, limit);

            var result = new
            {
                table = tableName,
                outputFile = outputPath,
                message = "Table exported successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (TableNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_tables to see available tables.";
        }
        catch (InvalidSqlException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool(Name = "export_access_query_to_csv")]
    [Description("Export Access query results to a CSV file")]
    public string ExportAccessQueryToCsv(
        [Description("Full file path to the Access database")]
        string filePath,
        [Description("Name of the query to execute")]
        string queryName,
        [Description("Full file path for the output CSV file")]
        string outputPath,
        [Description("Maximum number of rows to export (0 = unlimited, default: 0)")]
        int limit = 0)
    {
        try
        {
            _accessService.ExportQueryToCsv(filePath, queryName, outputPath, null, limit);

            var result = new
            {
                query = queryName,
                outputFile = outputPath,
                message = "Query results exported successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (QueryNotFoundException ex)
        {
            return $"Error: {ex.Message}. Use list_access_queries to see available queries.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    #endregion

    #region VBA Procedure Add/Delete Tools

    [McpServerTool(Name = "add_access_vba_procedure")]
    [Description("Add a new procedure to an Access VBA module. If a procedure with the same name exists, an error will be returned. Use write_access_vba_procedure if you want to replace an existing procedure.")]
    public string AddAccessVbaProcedure(
        [Description("Full file path to the Access database (e.g., C:\\MyDatabase.accdb)")] string filePath,
        [Description("Name of the VBA module to add the procedure to")] string moduleName,
        [Description("The complete VBA code for the procedure, including the procedure declaration (Sub/Function/Property) and End statement. The procedure name will be extracted from the code. IMPORTANT: Do NOT apply XML escaping (use '&' not '&amp;', '<' not '&lt;', '>' not '&gt;')")] string code,
        [Description("Insert the new procedure after this existing procedure (optional, default: append to end)")] string? insertAfter = null)
    {
        try
        {
            _accessService.AddProcedure(filePath, moduleName, code, insertAfter);

            var procedureName = CodeNormalizer.ExtractProcedureName(code);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                procedure = procedureName,
                insertedAfter = insertAfter,
                linesWritten = code.Split('\n').Length,
                message = "Procedure added successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Access's Trust Center settings.";
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
            return $"Error adding procedure: {ex.Message}";
        }
    }

    [McpServerTool(Name = "delete_access_vba_procedure")]
    [Description("Delete a procedure from an Access VBA module. IMPORTANT: This operation is irreversible.")]
    public string DeleteAccessVbaProcedure(
        [Description("Full file path to the Access database (e.g., C:\\MyDatabase.accdb)")] string filePath,
        [Description("Name of the VBA module containing the procedure")] string moduleName,
        [Description("Name of the procedure to delete (Sub, Function, or Property)")] string procedureName)
    {
        try
        {
            _accessService.DeleteProcedure(filePath, moduleName, procedureName);

            var result = new
            {
                success = true,
                file = filePath,
                module = moduleName,
                procedure = procedureName,
                deleted = true,
                message = "Procedure deleted successfully"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (FileNotFoundException)
        {
            return $"Error: Database not found or not open: {filePath}";
        }
        catch (VbaProjectAccessDeniedException ex)
        {
            return $"Error: {ex.Message}\n\nPlease enable 'Trust access to the VBA project object model' in Access's Trust Center settings.";
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
            return $"Error deleting procedure: {ex.Message}";
        }
    }

    #endregion
}
