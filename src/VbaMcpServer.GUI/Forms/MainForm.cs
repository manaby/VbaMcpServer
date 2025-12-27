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
    private string _mcpServerPath;
    private string? _selectedFilePath;
    private TargetFileInfo? _currentTargetFile;

    public MainForm()
    {
        InitializeComponent();

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

            // Check if file is selected
            if (string.IsNullOrEmpty(_selectedFilePath))
            {
                MessageBox.Show("Please select a target file first.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Open the file (if not already open)
            AppendServerLog($"Opening file: {_selectedFilePath}");
            _currentTargetFile = await _fileOpenerService.OpenFileAsync(_selectedFilePath);

            // Start file status monitoring
            _fileOpenerService.StartMonitoring(_selectedFilePath, TimeSpan.FromSeconds(5));

            // Start MCP server with target file
            _serverHost.Start(_mcpServerPath, null, _selectedFilePath);
            StartLogMonitoring();

            AppendServerLog($"Server started with target file: {_selectedFilePath}");

            // Update button states
            UpdateButtonStates();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to start server:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnStop_Click(object sender, EventArgs e)
    {
        try
        {
            _serverHost.Stop();
            _logViewer.StopWatching();

            // Stop file monitoring
            _fileOpenerService.StopMonitoring();

            // Update button states (re-enable Browse/Clear)
            UpdateButtonStates();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to stop server:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnRestart_Click(object sender, EventArgs e)
    {
        try
        {
            // Pass target file path to MCP server if selected
            var targetFile = _selectedFilePath;
            _serverHost.Restart(_mcpServerPath, null, targetFile);

            if (targetFile != null)
            {
                AppendServerLog($"Server restarted with target file: {targetFile}");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to restart server:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnClearLogs_Click(object sender, EventArgs e)
    {
        if (tabLogs.SelectedIndex == 0)
        {
            txtServerLog.Clear();
        }
        else
        {
            txtVbaLog.Clear();
        }
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
                var logs = tabLogs.SelectedIndex == 0
                    ? txtServerLog.Lines.ToList()
                    : txtVbaLog.Lines.ToList();

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

            var vbaLogs = _logViewer.LoadVbaEditLogs();
            txtVbaLog.Lines = vbaLogs.ToArray();
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
            Filter = "VBA Files|*.xlsm;*.xlsx;*.xlsb;*.xls;*.accdb;*.mdb|" +
                     "Excel Files|*.xlsm;*.xlsx;*.xlsb;*.xls|" +
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

            AppendServerLog($"Target file set: {filePath}");

            // Update button states (enable Start button)
            UpdateButtonStates();
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
        AppendServerLog("Target file cleared");

        // Update button states (disable Start button)
        UpdateButtonStates();
    }

    private void OnFileStatusChanged(object? sender, TargetFileInfo info)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnFileStatusChanged(sender, info)));
            return;
        }

        UpdateFileStatusLabel(info);
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
}
