using Microsoft.Extensions.Logging;
using VbaMcpServer.Exceptions;
using VbaMcpServer.Helpers;

namespace VbaMcpServer.Services;

/// <summary>
/// ExcelComServiceとAccessComServiceの共通基底クラス
/// COM参照管理とエラーハンドリングを集約
/// </summary>
/// <typeparam name="TApp">COM Applicationの型 (Excel.Application or Access.Application)</typeparam>
public abstract class ComServiceBase<TApp> where TApp : class
{
    protected readonly ILogger _logger;

    protected ComServiceBase(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// COM Application取得を実装（子クラスで実装）
    /// </summary>
    protected abstract TApp? GetApplication();

    /// <summary>
    /// アプリケーションが利用可能かチェック（子クラスで実装）
    /// </summary>
    protected abstract bool IsApplicationAvailable();

    /// <summary>
    /// COM操作を安全に実行（自動解放、エラーハンドリング）
    /// </summary>
    /// <typeparam name="TResult">戻り値の型</typeparam>
    /// <param name="operation">実行する操作（lambda）</param>
    /// <param name="operationName">操作名（ログ用）</param>
    /// <returns>操作の結果</returns>
    protected TResult ExecuteWithApp<TResult>(
        Func<TApp, TResult> operation,
        string operationName)
    {
        using var appWrapper = new ComObjectWrapper<TApp>(GetApplication(), _logger);
        var app = appWrapper.Value;

        if (app == null)
        {
            var appType = typeof(TApp).Name;
            _logger.LogError("{AppType} is not available for operation: {Operation}",
                appType, operationName);
            throw new ApplicationNotRunningException(
                $"{appType} is not running. Please open the application and try again.");
        }

        try
        {
            _logger.LogDebug("Executing {Operation}", operationName);
            var result = operation(app);
            _logger.LogDebug("Completed {Operation}", operationName);
            return result;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            _logger.LogError(ex, "COM error in {Operation}: 0x{HResult:X}",
                operationName, ex.HResult);
            throw new VbaOperationException(
                $"COM error in {operationName}: {ComErrorCodes.GetErrorMessage(ex.HResult)}",
                ex);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error in {Operation}", operationName);
            throw new VbaOperationException(
                $"Failed to execute {operationName}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// COM操作を安全に実行（void戻り値版）
    /// </summary>
    protected void ExecuteWithApp(
        Action<TApp> operation,
        string operationName)
    {
        ExecuteWithApp<object?>(app =>
        {
            operation(app);
            return null;
        }, operationName);
    }
}

/// <summary>
/// Access専用の基底クラス（CurrentDb()サポート）
/// </summary>
public abstract class AccessComServiceBase : ComServiceBase<dynamic>
{
    protected AccessComServiceBase(ILogger logger) : base(logger)
    {
    }

    /// <summary>
    /// CurrentDb()を使用するAccess操作を安全に実行
    /// </summary>
    /// <typeparam name="TResult">戻り値の型</typeparam>
    /// <param name="filePath">データベースファイルパス</param>
    /// <param name="operation">実行する操作（Access.Application, CurrentDb()を受け取る）</param>
    /// <param name="operationName">操作名（ログ用）</param>
    /// <returns>操作の結果</returns>
    protected TResult ExecuteWithDatabase<TResult>(
        string filePath,
        Func<dynamic, dynamic, TResult> operation,
        string operationName)
    {
        return ExecuteWithApp(app =>
        {
            // CurrentProjectの取得と検証
            dynamic? currentProject = null;
            try
            {
                currentProject = app.CurrentProject;
                var currentDbPath = currentProject?.FullName;

                if (!Path.GetFullPath(currentDbPath).Equals(
                    Path.GetFullPath(filePath),
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new FileNotFoundException(
                        $"Database '{filePath}' is not currently open in Access");
                }
            }
            finally
            {
                ReleaseComObject(currentProject);
            }

            // CurrentDb()の取得と操作実行
            using var dbWrapper = new ComObjectWrapper<dynamic>(app.CurrentDb(), _logger);
            var currentDb = dbWrapper.Value;

            return operation(app, currentDb);
        }, operationName);
    }

    /// <summary>
    /// CurrentDb()を使用するAccess操作を安全に実行（void版）
    /// </summary>
    protected void ExecuteWithDatabase(
        string filePath,
        Action<dynamic, dynamic> operation,
        string operationName)
    {
        ExecuteWithDatabase<object?>(filePath, (app, db) =>
        {
            operation(app, db);
            return null;
        }, operationName);
    }

    /// <summary>
    /// COMオブジェクトを安全に解放するヘルパー
    /// </summary>
    protected void ReleaseComObject(dynamic? obj)
    {
        if (obj != null)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to release COM object");
            }
        }
    }
}
