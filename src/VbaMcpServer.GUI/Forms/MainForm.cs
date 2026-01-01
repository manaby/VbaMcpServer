using VbaMcpServer.GUI.Models;
using VbaMcpServer.GUI.Services;
using VbaMcpServer.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;

namespace VbaMcpServer.GUI.Forms;

public partial class MainForm : Form
{
    private readonly McpServerHostService _serverHost;
    private readonly LogViewerService _logViewer;
    private readonly IConfiguration _configuration;
    private readonly FileOpenerService _fileOpenerService;
    private readonly GuiStateMachine _stateMachine;
    private string _mcpServerPath;
    private string? _selectedFilePath;
    private TargetFileInfo? _currentTargetFile;
    private CancellationTokenSource? _operationCts;

    public MainForm()
    {
        InitializeComponent();

        // Set window title with version
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        this.Text = $"VBA MCP Server Manager - v{version?.Major}.{version?.Minor}.{version?.Build}";

        // State Machine initialization
        _stateMachine = new GuiStateMachine();
        _stateMachine.StateChanged += OnGuiStateChanged;

        _serverHost = new McpServerHostService();
        _logViewer = new LogViewerService();

        // Load configuration
        _configuration = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .Build();

        // Find MCP server executable
        _mcpServerPath = FindMcpServerExecutable();

        // Initialize FileOpenerService
        var excelComService = new ExcelComService(
            LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<ExcelComService>()
        );
        _fileOpenerService = new FileOpenerService(
            LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<FileOpenerService>(),
            excelComService
        );
        _fileOpenerService.FileStatusChanged += OnFileStatusChanged;

        // Subscribe to server events
        _serverHost.StatusChanged += OnServerStatusChanged;
        _serverHost.OutputReceived += OnServerOutputReceived;
        _serverHost.ErrorReceived += OnServerErrorReceived;

        // Initialize status
        UpdateStatusDisplay(ServerStatus.Stopped, null);

        // Initialize button states
        UpdateButtonStates();

        // Log the found path for debugging
        AppendServerLog($"MCP Server path: {_mcpServerPath}");
        AppendServerLog($"Current directory: {AppDomain.CurrentDomain.BaseDirectory}");

        // Load existing logs
        LoadLogs();

        // Initial state is already Idle_NoFile (default in GuiStateMachine)
        // Manually trigger initial button state update
        UpdateButtonStatesForState(GuiState.Idle_NoFile);
        UpdateStatusForState(GuiState.Idle_NoFile);
        AppendServerLog("State: Application started (Idle_NoFile)");
    }

    private string FindMcpServerExecutable()
    {
        var currentDir = AppDomain.CurrentDomain.BaseDirectory;
        var candidates = new List<string>();

        AppendServerLog($"Current directory: {currentDir}");

        // Priority 1: User configuration in appsettings.json
        var configPath = _configuration["VbaMcpServer:ServerExePath"];
        if (!string.IsNullOrWhiteSpace(configPath))
        {
            candidates.Add(configPath);
            AppendServerLog($"Found path in appsettings.json: {configPath}");
        }

        // Priority 2: Registry (set by installer)
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\VbaMcpServer");
            var registryPath = key?.GetValue("ServerExePath") as string;
            if (!string.IsNullOrWhiteSpace(registryPath))
            {
                candidates.Add(registryPath);
                AppendServerLog($"Found path in registry: {registryPath}");
            }
        }
        catch (Exception ex)
        {
            AppendServerLog($"Registry read failed (this is normal for development): {ex.Message}");
        }

        // Priority 3: Same directory (unified output directory for development and installed location)
        var sameDirPath = Path.Combine(currentDir, "VbaMcpServer.exe");
        candidates.Add(sameDirPath);
        AppendServerLog($"Checking same directory: {sameDirPath}");

        // Search through all candidates
        foreach (var candidate in candidates)
        {
            try
            {
                var fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                {
                    AppendServerLog($"✓ Found VbaMcpServer.exe: {fullPath}");
                    return fullPath;
                }
                else
                {
                    AppendServerLog($"✗ Not found: {fullPath}");
                }
            }
            catch (Exception ex)
            {
                AppendServerLog($"Error checking path '{candidate}': {ex.Message}");
            }
        }

        AppendServerLog("========================================");
        AppendServerLog("ERROR: VbaMcpServer.exe not found");
        AppendServerLog("========================================");
        AppendServerLog("Solutions:");
        AppendServerLog("1. Build the entire solution (Build > Build Solution)");
        AppendServerLog("2. Or specify custom path in appsettings.json:");
        AppendServerLog("   { \"VbaMcpServer\": { \"ServerExePath\": \"C:\\\\path\\\\to\\\\VbaMcpServer.exe\" } }");
        AppendServerLog("========================================");

        return Path.Combine(currentDir, "VbaMcpServer.exe"); // Default fallback
    }

    private async void btnStart_Click(object sender, EventArgs e)
    {
        try
        {
            if (!File.Exists(_mcpServerPath))
            {
                MessageBox.Show($"MCP Server executable not found:\n{_mcpServerPath}\n\nPlease build the VbaMcpServer project first.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(_selectedFilePath))
            {
                MessageBox.Show("Please select a target file first.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 新しいCancellationTokenSourceを作成
            _operationCts?.Dispose();
            _operationCts = new CancellationTokenSource();

            await StartServerAsync(_selectedFilePath, _operationCts.Token);
        }
        catch (OperationCanceledException)
        {
            AppendServerLog("Start operation cancelled by user");
            _stateMachine.TransitionTo(GuiState.Idle_FileSelected, "Cancelled");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to start server:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _stateMachine.TransitionTo(GuiState.Error_FileOpenFailed, ex.Message);
        }
        finally
        {
            _operationCts?.Dispose();
            _operationCts = null;
        }
    }

    /// <summary>
    /// サーバ起動の非同期処理（State Machine統合版）
    /// </summary>
    private async Task StartServerAsync(string filePath, CancellationToken cancellationToken)
    {
        // Step 1: ファイルを開く
        _stateMachine.TransitionTo(GuiState.Starting_OpeningFile,
            $"Opening {Path.GetFileName(filePath)}");

        TargetFileInfo fileInfo;
        try
        {
            fileInfo = await _fileOpenerService.OpenFileAsync(filePath);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (Exception ex)
        {
            _stateMachine.TransitionTo(GuiState.Error_FileOpenFailed, ex.Message);
            throw;
        }

        // Step 2: ファイルが開くまで待機
        _stateMachine.TransitionTo(GuiState.Starting_WaitingForFile,
            "Waiting for file to open");

        if (!fileInfo.IsOpen)
        {
            _stateMachine.TransitionTo(GuiState.Error_FileOpenFailed,
                "File did not open within timeout");
            throw new TimeoutException("File did not open within timeout");
        }

        _currentTargetFile = fileInfo;

        // Step 3: ファイル監視を開始
        _fileOpenerService.StartMonitoring(filePath, TimeSpan.FromSeconds(5));

        // Step 4: MCPサーバ起動
        _stateMachine.TransitionTo(GuiState.Starting_LaunchingServer,
            "Launching MCP server");

        cancellationToken.ThrowIfCancellationRequested();

        // Task.Run()でサーバ起動をバックグラウンドスレッドで実行
        await Task.Run(() =>
        {
            _serverHost.Start(_mcpServerPath, null, filePath);
        }, cancellationToken).ConfigureAwait(false);

        StartLogMonitoring();

        // Step 5: 実行中状態に遷移
        _stateMachine.TransitionTo(GuiState.Running_FileOpen,
            $"Server started with PID {_serverHost.ProcessId}");
    }

    private async void btnStop_Click(object sender, EventArgs e)
    {
        // Cancel any ongoing operation
        _operationCts?.Cancel();

        await StopServerAsync();
    }

    /// <summary>
    /// Stop the MCP server asynchronously with state transitions
    /// </summary>
    private async Task StopServerAsync()
    {
        try
        {
            // Phase 1: Server shutdown
            _stateMachine.TransitionTo(GuiState.Stopping_ServerShutdown,
                "Stopping MCP server...");

            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await _serverHost.StopAsync(cts.Token);

            _logViewer.StopWatching();

            // Phase 2: Cleanup
            _stateMachine.TransitionTo(GuiState.Stopping_Cleanup,
                "Cleaning up resources...");

            // Stop file monitoring
            _fileOpenerService.StopMonitoring();

            // Brief delay to show cleanup state
            await Task.Delay(100);

            // Transition back to Idle_FileSelected
            _stateMachine.TransitionTo(GuiState.Idle_FileSelected,
                "Server stopped successfully");
        }
        catch (Exception ex)
        {
            AppendServerLog($"[ERROR] Failed to stop server: {ex.Message}");
            MessageBox.Show($"Failed to stop server:\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            // On error, force cleanup and return to Idle_FileSelected
            _fileOpenerService.StopMonitoring();
            _logViewer.StopWatching();
            _stateMachine.TransitionTo(GuiState.Idle_FileSelected,
                "Server stopped with errors");
        }
    }

    private async void btnRestart_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_selectedFilePath))
        {
            MessageBox.Show("No file selected for restart.",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            // Step 1: Stop the server
            await StopServerAsync();

            // Step 2: Start the server with the same file
            _operationCts?.Dispose();
            _operationCts = new CancellationTokenSource();
            await StartServerAsync(_selectedFilePath, _operationCts.Token);
        }
        catch (Exception ex)
        {
            AppendServerLog($"[ERROR] Failed to restart server: {ex.Message}");
            MessageBox.Show($"Failed to restart server:\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnClearLogs_Click(object sender, EventArgs e)
    {
        txtServerLog.Clear();
    }

    private void btnSaveLogs_Click(object sender, EventArgs e)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
            DefaultExt = "txt",
            FileName = $"mcp-server-log-{DateTime.Now:yyyy-MM-dd-HHmmss}.txt"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var logs = txtServerLog.Lines.ToList();

                _logViewer.SaveLogsToFile(dialog.FileName, logs);
                MessageBox.Show($"Logs saved to:\n{dialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save logs:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void OnServerStatusChanged(object? sender, ServerStatusChangedEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnServerStatusChanged(sender, e)));
            return;
        }

        UpdateStatusDisplay(e.Status, e.ProcessId);

        if (!string.IsNullOrEmpty(e.Message))
        {
            AppendServerLog(e.Message);
        }

        // State Machine integration: Detect server crash
        var currentState = _stateMachine.CurrentState;
        if (e.Status == ServerStatus.Stopped &&
            (currentState == GuiState.Running_FileOpen ||
             currentState == GuiState.Running_FileClosedByUser))
        {
            // Server crashed unexpectedly
            _stateMachine.TransitionTo(GuiState.Error_ServerCrashed,
                "Server process crashed unexpectedly.");

            // Stop file monitoring on crash
            _fileOpenerService.StopMonitoring();
            _logViewer.StopWatching();
        }
    }

    private void OnServerOutputReceived(object? sender, string message)
    {
        AppendServerLog($"[OUT] {message}");
    }

    private void OnServerErrorReceived(object? sender, string message)
    {
        // Check if the message already contains a log level indicator
        // Serilog often writes all logs to stderr, including [INF], [WRN], [ERR]
        if (message.Contains("[INF]") || message.Contains("[WRN]") || message.Contains("[ERR]") ||
            message.Contains("[DBG]") || message.Contains("[VRB]") || message.Contains("[FTL]"))
        {
            // Message already has log level, don't add [ERR] prefix
            AppendServerLog(message);
        }
        else
        {
            // No log level found, this is a genuine error message
            AppendServerLog($"[ERR] {message}");
        }
    }

    private void UpdateStatusDisplay(ServerStatus status, int? processId)
    {
        lblStatus.Text = $"Status: {status}";
        lblProcessId.Text = processId.HasValue ? $"Process ID: {processId}" : "Process ID: N/A";

        switch (status)
        {
            case ServerStatus.Stopped:
                lblStatus.ForeColor = Color.Red;
                break;
            case ServerStatus.Starting:
                lblStatus.ForeColor = Color.Orange;
                break;
            case ServerStatus.Running:
                lblStatus.ForeColor = Color.Green;
                break;
            case ServerStatus.Stopping:
                lblStatus.ForeColor = Color.Orange;
                break;
        }

        // Update button states based on server status and file selection
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        bool serverRunning = _serverHost.CurrentStatus == ServerStatus.Running ||
                            _serverHost.CurrentStatus == ServerStatus.Starting ||
                            _serverHost.CurrentStatus == ServerStatus.Stopping;
        bool fileSelected = !string.IsNullOrEmpty(_selectedFilePath);

        // File selection controls - only enabled when server is stopped
        btnBrowseFile.Enabled = !serverRunning;
        btnClearFile.Enabled = !serverRunning && fileSelected;

        // Server control buttons
        btnStart.Enabled = !serverRunning && fileSelected;
        btnStop.Enabled = _serverHost.CurrentStatus == ServerStatus.Running;
        btnRestart.Enabled = _serverHost.CurrentStatus == ServerStatus.Running;
    }

    private void AppendServerLog(string message)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => AppendServerLog(message)));
            return;
        }

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        txtServerLog.AppendText($"{timestamp} - {message}{Environment.NewLine}");
        txtServerLog.SelectionStart = txtServerLog.Text.Length;
        txtServerLog.ScrollToCaret();
    }

    private void LoadLogs()
    {
        try
        {
            var serverLogs = _logViewer.LoadServerLogs();
            txtServerLog.Lines = serverLogs.ToArray();
        }
        catch (Exception ex)
        {
            AppendServerLog($"Error loading logs: {ex.Message}");
        }
    }

    private void StartLogMonitoring()
    {
        _logViewer.StartWatching("server", line =>
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() =>
                {
                    txtServerLog.AppendText(line + Environment.NewLine);
                    txtServerLog.SelectionStart = txtServerLog.Text.Length;
                    txtServerLog.ScrollToCaret();
                }));
            }
        });
    }

    private void btnBrowseFile_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Select VBA File",
            Filter = "VBA Files|*.xlsm;*.xlsb;*.xls;*.accdb;*.mdb|" +
                     "Excel Macro Files|*.xlsm;*.xlsb;*.xls|" +
                     "Access Files|*.accdb;*.mdb|" +
                     "All Files|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            var filePath = dialog.FileName;
            AppendServerLog($"Selected file: {filePath}");

            // Store file path only (do not open the file)
            _selectedFilePath = filePath;
            txtFilePath.Text = filePath;
            lblFileStatus.Text = "Status: File selected (not opened)";
            lblFileStatus.ForeColor = Color.Blue;

            // State transition - only transition if not already in Idle_FileSelected
            if (_stateMachine.CurrentState != GuiState.Idle_FileSelected)
            {
                _stateMachine.TransitionTo(GuiState.Idle_FileSelected,
                    $"File selected: {Path.GetFileName(filePath)}");
            }
            else
            {
                // Already in the correct state, just log the change
                AppendServerLog($"File changed: {Path.GetFileName(filePath)}");
            }
        }
    }

    private void btnClearFile_Click(object sender, EventArgs e)
    {
        _fileOpenerService.StopMonitoring();
        _selectedFilePath = null;
        _currentTargetFile = null;
        txtFilePath.Text = "(Select a file)";
        lblFileStatus.Text = "Status: Not selected";
        lblFileStatus.ForeColor = Color.Gray;

        // State transition
        _stateMachine.TransitionTo(GuiState.Idle_NoFile, "File cleared");
    }

    private void OnFileStatusChanged(object? sender, TargetFileInfo info)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnFileStatusChanged(sender, info)));
            return;
        }

        UpdateFileStatusLabel(info);

        // State Machine integration
        var currentState = _stateMachine.CurrentState;

        // If server is running and file was closed by user
        if (currentState == GuiState.Running_FileOpen && !info.IsOpen)
        {
            _stateMachine.TransitionTo(GuiState.Running_FileClosedByUser,
                "Warning: File was closed. Please keep the file open while server is running.");
        }
        // If server is running with file closed warning and file was reopened
        else if (currentState == GuiState.Running_FileClosedByUser && info.IsOpen)
        {
            _stateMachine.TransitionTo(GuiState.Running_FileOpen,
                "File reopened successfully.");
        }
    }

    private void UpdateFileStatusLabel(TargetFileInfo info)
    {
        var appName = info.FileType == FileType.Excel ? "Excel" : "Access";

        if (info.IsOpen && info.ProcessId.HasValue)
        {
            lblFileStatus.Text = $"Status: ● Opened in {appName} (PID: {info.ProcessId})";
            lblFileStatus.ForeColor = Color.Green;
        }
        else if (info.IsOpen)
        {
            lblFileStatus.Text = $"Status: ● Opened in {appName}";
            lblFileStatus.ForeColor = Color.Green;
        }
        else
        {
            lblFileStatus.Text = "Status: ○ File is closed";
            lblFileStatus.ForeColor = Color.Orange;
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_serverHost.CurrentStatus == ServerStatus.Running)
        {
            var result = MessageBox.Show(
                "MCP Server is still running. Do you want to stop it before closing?",
                "Confirm Exit",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == DialogResult.Yes)
            {
                _serverHost.Stop();
            }
        }

        _fileOpenerService.Dispose();
        _logViewer.Dispose();
        _serverHost.Dispose();

        base.OnFormClosing(e);
    }

    #region State Machine Event Handlers

    /// <summary>
    /// GUI状態変化時のイベントハンドラ（UIスレッドで呼ばれる前提）
    /// </summary>
    private void OnGuiStateChanged(object? sender, StateChangedEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnGuiStateChanged(sender, e)));
            return;
        }

        // ログ記録
        AppendServerLog($"State: {e.PreviousState} → {e.NewState}" +
            (e.Message != null ? $" ({e.Message})" : ""));

        // ボタン状態を更新
        UpdateButtonStatesForState(e.NewState);

        // ステータス表示を更新
        UpdateStatusForState(e.NewState);
    }

    /// <summary>
    /// 状態に基づいてボタンの有効/無効を設定
    /// </summary>
    private void UpdateButtonStatesForState(GuiState state)
    {
        switch (state)
        {
            case GuiState.Idle_NoFile:
                btnBrowseFile.Enabled = true;
                btnClearFile.Enabled = false;
                btnStart.Enabled = false;
                btnStop.Enabled = false;
                btnRestart.Enabled = false;
                btnForceStop.Visible = false;
                progressBar.Visible = false;
                pnlWarningBanner.Visible = false;
                break;

            case GuiState.Idle_FileSelected:
                btnBrowseFile.Enabled = true;
                btnClearFile.Enabled = true;
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                btnRestart.Enabled = false;
                btnForceStop.Visible = false;
                progressBar.Visible = false;
                pnlWarningBanner.Visible = false;
                break;

            case GuiState.Starting_OpeningFile:
            case GuiState.Starting_WaitingForFile:
            case GuiState.Starting_LaunchingServer:
                btnBrowseFile.Enabled = false;
                btnClearFile.Enabled = false;
                btnStart.Enabled = false;
                btnStop.Enabled = true;  // Cancel可能
                btnRestart.Enabled = false;
                btnForceStop.Visible = false;
                progressBar.Visible = true;  // プログレスバー表示
                pnlWarningBanner.Visible = false;
                break;

            case GuiState.Running_FileOpen:
                btnBrowseFile.Enabled = false;
                btnClearFile.Enabled = false;
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                btnRestart.Enabled = true;
                btnForceStop.Visible = false;
                progressBar.Visible = false;
                pnlWarningBanner.Visible = false;  // 警告非表示
                break;

            case GuiState.Running_FileClosedByUser:
                btnBrowseFile.Enabled = false;
                btnClearFile.Enabled = false;
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                btnRestart.Enabled = true;
                btnForceStop.Visible = false;
                progressBar.Visible = false;
                pnlWarningBanner.Visible = true;  // 警告バナー表示
                break;

            case GuiState.Stopping_ServerShutdown:
            case GuiState.Stopping_Cleanup:
                btnBrowseFile.Enabled = false;
                btnClearFile.Enabled = false;
                btnStart.Enabled = false;
                btnStop.Enabled = false;
                btnRestart.Enabled = false;
                btnForceStop.Visible = false;
                progressBar.Visible = true;  // プログレスバー表示
                pnlWarningBanner.Visible = false;
                break;

            case GuiState.Error_FileOpenFailed:
            case GuiState.Error_ServerCrashed:
                btnBrowseFile.Enabled = true;
                btnClearFile.Enabled = true;
                btnStart.Enabled = true;  // Retry可能
                btnStop.Enabled = false;
                btnRestart.Enabled = false;
                btnForceStop.Visible = false;
                progressBar.Visible = false;
                pnlWarningBanner.Visible = false;
                break;
        }
    }

    /// <summary>
    /// 状態に基づいてステータス表示を更新
    /// </summary>
    private void UpdateStatusForState(GuiState state)
    {
        switch (state)
        {
            case GuiState.Idle_NoFile:
            case GuiState.Idle_FileSelected:
                lblStatus.Text = "Status: Stopped";
                lblStatus.ForeColor = Color.Gray;
                break;

            case GuiState.Starting_OpeningFile:
                lblStatus.Text = "Status: Opening file...";
                lblStatus.ForeColor = Color.Orange;
                break;

            case GuiState.Starting_WaitingForFile:
                lblStatus.Text = "Status: Waiting for file...";
                lblStatus.ForeColor = Color.Orange;
                break;

            case GuiState.Starting_LaunchingServer:
                lblStatus.Text = "Status: Launching server...";
                lblStatus.ForeColor = Color.Orange;
                break;

            case GuiState.Running_FileOpen:
                lblStatus.Text = "Status: Running";
                lblStatus.ForeColor = Color.Green;
                break;

            case GuiState.Running_FileClosedByUser:
                lblStatus.Text = "Status: Running (⚠ File closed)";
                lblStatus.ForeColor = Color.DarkOrange;
                break;

            case GuiState.Stopping_ServerShutdown:
            case GuiState.Stopping_Cleanup:
                lblStatus.Text = "Status: Stopping...";
                lblStatus.ForeColor = Color.Orange;
                break;

            case GuiState.Error_FileOpenFailed:
                lblStatus.Text = "Status: Error - File open failed";
                lblStatus.ForeColor = Color.Red;
                break;

            case GuiState.Error_ServerCrashed:
                lblStatus.Text = "Status: Error - Server crashed";
                lblStatus.ForeColor = Color.Red;
                break;
        }
    }

    #endregion
}
