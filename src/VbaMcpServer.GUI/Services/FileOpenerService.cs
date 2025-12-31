using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using VbaMcpServer.Services;
using VbaMcpServer.Helpers;
using VbaMcpServer.GUI.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;

namespace VbaMcpServer.GUI.Services;

/// <summary>
/// Excel/Access ファイルの開閉を管理するサービス
/// </summary>
public class FileOpenerService : IDisposable
{
    private readonly ILogger<FileOpenerService> _logger;
    private readonly ExcelComService _excelComService;
    private System.Windows.Forms.Timer? _statusCheckTimer;

    public event EventHandler<TargetFileInfo>? FileStatusChanged;

    public FileOpenerService(
        ILogger<FileOpenerService> logger,
        ExcelComService excelComService)
    {
        _logger = logger;
        _excelComService = excelComService;
    }

    /// <summary>
    /// ファイル種別を判定
    /// </summary>
    public FileType DetectFileType(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".xlsm" or ".xlsx" or ".xlsb" or ".xls" => FileType.Excel,
            ".accdb" or ".mdb" => FileType.Access,
            _ => FileType.Unknown
        };
    }

    /// <summary>
    /// ファイルが既に開いているか確認
    /// </summary>
    public TargetFileInfo CheckFileStatus(string filePath)
    {
        var fileType = DetectFileType(filePath);
        var info = new TargetFileInfo
        {
            FilePath = filePath,
            FileType = fileType,
            IsOpen = false,
            ProcessId = null,
            LastChecked = DateTime.Now
        };

        try
        {
            if (fileType == FileType.Excel)
            {
                // ExcelComService を使用して確認
                Excel.Workbook? workbook = null;
                try
                {
                    workbook = _excelComService.GetWorkbook(filePath);
                    if (workbook != null)
                    {
                        info.IsOpen = true;
                        // プロセスIDの取得（Excel.Application から）
                        Excel.Application? excelApp = null;
                        try
                        {
                            excelApp = (Excel.Application)ComHelper.GetActiveObject("Excel.Application");
                            info.ProcessId = GetProcessIdFromHwnd(excelApp.Hwnd);
                            _logger.LogInformation("File is already open in Excel: {FilePath}", filePath);
                        }
                        catch (COMException ex)
                        {
                            _logger.LogWarning(ex, "Could not get Excel process ID");
                        }
                        finally
                        {
                            // CRITICAL: Release the COM object to prevent memory leak
                            if (excelApp != null)
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(excelApp);
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogWarning(ex, "Failed to release Excel COM object");
                                }
                                excelApp = null;
                            }
                        }
                    }
                }
                finally
                {
                    // CRITICAL: Release the workbook COM object
                    if (workbook != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(workbook);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to release Workbook COM object");
                        }
                        workbook = null;
                    }
                }
            }
            else if (fileType == FileType.Access)
            {
                // Access の場合も同様
                Access.Application? accessApp = null;
                dynamic? currentProject = null;
                try
                {
                    accessApp = (Access.Application)ComHelper.GetActiveObject("Access.Application");

                    // CRITICAL: Store CurrentProject in variable to release RCW later
                    currentProject = accessApp.CurrentProject;
                    if (currentProject.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        info.IsOpen = true;
                        info.ProcessId = GetProcessIdFromWindow(accessApp.hWndAccessApp());
                        _logger.LogInformation("File is already open in Access: {FilePath}", filePath);
                    }
                }
                catch (COMException)
                {
                    // Access が起動していない
                }
                finally
                {
                    // CRITICAL: Release CurrentProject RCW first (fixes root cause leak)
                    if (currentProject != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(currentProject);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to release CurrentProject COM object");
                        }
                        currentProject = null;
                    }

                    // CRITICAL: Release the Access.Application COM object
                    if (accessApp != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(accessApp);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to release Access COM object");
                        }
                        accessApp = null;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to check file status: {FilePath}", filePath);
        }

        return info;
    }

    /// <summary>
    /// ファイルを開く（既に開いていない場合）
    /// </summary>
    public async Task<TargetFileInfo> OpenFileAsync(string filePath)
    {
        var status = CheckFileStatus(filePath);

        if (status.IsOpen)
        {
            _logger.LogInformation("File is already open: {FilePath}", filePath);
            return status;
        }

        // ファイルを開く
        try
        {
            var fileType = status.FileType;
            Process? process = null;

            if (fileType == FileType.Excel)
            {
                _logger.LogInformation("Opening Excel file: {FilePath}", filePath);
                process = Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            else if (fileType == FileType.Access)
            {
                _logger.LogInformation("Opening Access file: {FilePath}", filePath);
                process = Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            else
            {
                throw new InvalidOperationException($"Unsupported file type: {Path.GetExtension(filePath)}");
            }

            // アプリケーションが完全に起動するまで待機
            if (process != null)
            {
                await Task.Delay(3000); // 3秒待機
                status = CheckFileStatus(filePath);

                // 最大10秒まで待機
                int maxRetries = 10;
                while (!status.IsOpen && maxRetries > 0)
                {
                    await Task.Delay(1000);
                    status = CheckFileStatus(filePath);
                    maxRetries--;
                }

                if (status.IsOpen)
                {
                    _logger.LogInformation("File opened successfully: {FilePath}", filePath);
                }
                else
                {
                    _logger.LogWarning("File may not have opened correctly: {FilePath}", filePath);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to open file: {FilePath}", filePath);
            throw;
        }

        return status;
    }

    /// <summary>
    /// ファイル状態の定期監視を開始
    /// </summary>
    public void StartMonitoring(string filePath, TimeSpan interval)
    {
        StopMonitoring();

        _statusCheckTimer = new System.Windows.Forms.Timer
        {
            Interval = (int)interval.TotalMilliseconds
        };

        _statusCheckTimer.Tick += (sender, e) =>
        {
            var status = CheckFileStatus(filePath);
            FileStatusChanged?.Invoke(this, status);
        };

        _statusCheckTimer.Start();
        _logger.LogInformation("Started monitoring file status: {FilePath}", filePath);
    }

    /// <summary>
    /// ファイル状態の監視を停止
    /// </summary>
    public void StopMonitoring()
    {
        if (_statusCheckTimer != null)
        {
            _statusCheckTimer.Stop();
            _statusCheckTimer.Dispose();
            _statusCheckTimer = null;
            _logger.LogInformation("Stopped monitoring file status");
        }
    }

    private int GetProcessIdFromHwnd(int hwnd)
    {
        GetWindowThreadProcessId(hwnd, out uint processId);
        return (int)processId;
    }

    private int GetProcessIdFromWindow(int hwnd)
    {
        GetWindowThreadProcessId(hwnd, out uint processId);
        return (int)processId;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(int hWnd, out uint lpdwProcessId);

    public void Dispose()
    {
        StopMonitoring();
    }
}
