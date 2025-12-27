using System.Diagnostics;
using VbaMcpServer.GUI.Models;

namespace VbaMcpServer.GUI.Services;

public class McpServerHostService : IDisposable
{
    private Process? _serverProcess;
    private System.Windows.Forms.Timer? _monitorTimer;
    private readonly object _lock = new();

    public event EventHandler<ServerStatusChangedEventArgs>? StatusChanged;
    public event EventHandler<string>? OutputReceived;
    public event EventHandler<string>? ErrorReceived;

    public ServerStatus CurrentStatus { get; private set; } = ServerStatus.Stopped;
    public int? ProcessId => _serverProcess?.Id;

    public void Start(string exePath, string? workingDirectory = null, string? targetFilePath = null)
    {
        lock (_lock)
        {
            if (CurrentStatus == ServerStatus.Running || CurrentStatus == ServerStatus.Starting)
            {
                throw new InvalidOperationException("Server is already running or starting");
            }

            if (!File.Exists(exePath))
            {
                throw new FileNotFoundException($"MCP Server executable not found: {exePath}", exePath);
            }

            UpdateStatus(ServerStatus.Starting, null, "Starting MCP server...");

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    CreateNoWindow = true,
                    WorkingDirectory = workingDirectory ?? Path.GetDirectoryName(exePath) ?? Environment.CurrentDirectory
                };

                // Pass target file path as environment variable
                if (!string.IsNullOrEmpty(targetFilePath))
                {
                    startInfo.EnvironmentVariables["VBA_TARGET_FILE"] = targetFilePath;
                }

                _serverProcess = new Process { StartInfo = startInfo };
                _serverProcess.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        OutputReceived?.Invoke(this, e.Data);
                    }
                };
                _serverProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        ErrorReceived?.Invoke(this, e.Data);
                    }
                };
                _serverProcess.Exited += (sender, e) =>
                {
                    UpdateStatus(ServerStatus.Stopped, null, "Server process exited");
                };

                _serverProcess.EnableRaisingEvents = true;
                _serverProcess.Start();
                _serverProcess.BeginOutputReadLine();
                _serverProcess.BeginErrorReadLine();

                StartMonitoring();
                UpdateStatus(ServerStatus.Running, _serverProcess.Id, "MCP server started successfully");
            }
            catch (Exception ex)
            {
                UpdateStatus(ServerStatus.Stopped, null, $"Failed to start server: {ex.Message}");
                throw;
            }
        }
    }

    public void Stop()
    {
        lock (_lock)
        {
            if (CurrentStatus == ServerStatus.Stopped || CurrentStatus == ServerStatus.Stopping)
            {
                return;
            }

            UpdateStatus(ServerStatus.Stopping, ProcessId, "Stopping MCP server...");
            StopMonitoring();

            try
            {
                if (_serverProcess != null && !_serverProcess.HasExited)
                {
                    _serverProcess.StandardInput.Close();

                    if (!_serverProcess.WaitForExit(5000))
                    {
                        _serverProcess.Kill();
                        _serverProcess.WaitForExit();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorReceived?.Invoke(this, $"Error stopping server: {ex.Message}");
            }
            finally
            {
                _serverProcess?.Dispose();
                _serverProcess = null;
                UpdateStatus(ServerStatus.Stopped, null, "MCP server stopped");
            }
        }
    }

    public void Restart(string exePath, string? workingDirectory = null, string? targetFilePath = null)
    {
        Stop();
        Thread.Sleep(1000);
        Start(exePath, workingDirectory, targetFilePath);
    }

    private void StartMonitoring()
    {
        _monitorTimer = new System.Windows.Forms.Timer
        {
            Interval = 1000
        };
        _monitorTimer.Tick += MonitorTimer_Tick;
        _monitorTimer.Start();
    }

    private void StopMonitoring()
    {
        if (_monitorTimer != null)
        {
            _monitorTimer.Stop();
            _monitorTimer.Dispose();
            _monitorTimer = null;
        }
    }

    private void MonitorTimer_Tick(object? sender, EventArgs e)
    {
        lock (_lock)
        {
            if (_serverProcess != null && _serverProcess.HasExited && CurrentStatus == ServerStatus.Running)
            {
                UpdateStatus(ServerStatus.Stopped, null,
                    $"Server process exited unexpectedly (Exit code: {_serverProcess.ExitCode})");
            }
        }
    }

    private void UpdateStatus(ServerStatus newStatus, int? processId, string? message = null)
    {
        CurrentStatus = newStatus;
        StatusChanged?.Invoke(this, new ServerStatusChangedEventArgs(newStatus, processId, message));
    }

    public void Dispose()
    {
        StopMonitoring();
        Stop();
    }
}
