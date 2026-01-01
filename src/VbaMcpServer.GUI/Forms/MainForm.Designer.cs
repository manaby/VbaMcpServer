namespace VbaMcpServer.GUI.Forms;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        this.components = new System.ComponentModel.Container();
        this.grpServerControl = new System.Windows.Forms.GroupBox();
        this.lblStatus = new System.Windows.Forms.Label();
        this.lblProcessId = new System.Windows.Forms.Label();
        this.btnStart = new System.Windows.Forms.Button();
        this.btnStop = new System.Windows.Forms.Button();
        this.btnRestart = new System.Windows.Forms.Button();
        this.btnForceStop = new System.Windows.Forms.Button();
        this.progressBar = new System.Windows.Forms.ProgressBar();
        this.grpTargetFile = new System.Windows.Forms.GroupBox();
        this.txtFilePath = new System.Windows.Forms.TextBox();
        this.btnBrowseFile = new System.Windows.Forms.Button();
        this.lblFileStatus = new System.Windows.Forms.Label();
        this.btnClearFile = new System.Windows.Forms.Button();
        this.pnlWarningBanner = new System.Windows.Forms.Panel();
        this.lblWarningMessage = new System.Windows.Forms.Label();
        this.lblBackupWarning = new System.Windows.Forms.Label();
        this.grpLogs = new System.Windows.Forms.GroupBox();
        this.tabLogs = new System.Windows.Forms.TabControl();
        this.tabPageServerLog = new System.Windows.Forms.TabPage();
        this.txtServerLog = new System.Windows.Forms.TextBox();
        this.btnClearLogs = new System.Windows.Forms.Button();
        this.btnSaveLogs = new System.Windows.Forms.Button();
        this.grpServerControl.SuspendLayout();
        this.grpTargetFile.SuspendLayout();
        this.grpLogs.SuspendLayout();
        this.tabLogs.SuspendLayout();
        this.tabPageServerLog.SuspendLayout();
        this.SuspendLayout();

        // grpServerControl
        this.grpServerControl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.grpServerControl.Controls.Add(this.lblStatus);
        this.grpServerControl.Controls.Add(this.lblProcessId);
        this.grpServerControl.Controls.Add(this.btnStart);
        this.grpServerControl.Controls.Add(this.btnStop);
        this.grpServerControl.Controls.Add(this.btnRestart);
        this.grpServerControl.Controls.Add(this.btnForceStop);
        this.grpServerControl.Controls.Add(this.progressBar);
        this.grpServerControl.Location = new System.Drawing.Point(12, 148);
        this.grpServerControl.Name = "grpServerControl";
        this.grpServerControl.Size = new System.Drawing.Size(760, 100);
        this.grpServerControl.TabIndex = 0;
        this.grpServerControl.TabStop = false;
        this.grpServerControl.Text = "Server Control";

        // lblStatus
        this.lblStatus.AutoSize = true;
        this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
        this.lblStatus.Location = new System.Drawing.Point(15, 25);
        this.lblStatus.Name = "lblStatus";
        this.lblStatus.Size = new System.Drawing.Size(100, 19);
        this.lblStatus.TabIndex = 0;
        this.lblStatus.Text = "Status: Stopped";

        // lblProcessId
        this.lblProcessId.AutoSize = true;
        this.lblProcessId.Location = new System.Drawing.Point(15, 50);
        this.lblProcessId.Name = "lblProcessId";
        this.lblProcessId.Size = new System.Drawing.Size(95, 15);
        this.lblProcessId.TabIndex = 1;
        this.lblProcessId.Text = "Process ID: N/A";

        // btnStart
        this.btnStart.Location = new System.Drawing.Point(250, 25);
        this.btnStart.Name = "btnStart";
        this.btnStart.Size = new System.Drawing.Size(100, 30);
        this.btnStart.TabIndex = 2;
        this.btnStart.Text = "Start";
        this.btnStart.UseVisualStyleBackColor = true;
        this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

        // btnStop
        this.btnStop.Enabled = false;
        this.btnStop.Location = new System.Drawing.Point(360, 25);
        this.btnStop.Name = "btnStop";
        this.btnStop.Size = new System.Drawing.Size(100, 30);
        this.btnStop.TabIndex = 3;
        this.btnStop.Text = "Stop";
        this.btnStop.UseVisualStyleBackColor = true;
        this.btnStop.Click += new System.EventHandler(this.btnStop_Click);

        // btnRestart
        this.btnRestart.Enabled = false;
        this.btnRestart.Location = new System.Drawing.Point(470, 25);
        this.btnRestart.Name = "btnRestart";
        this.btnRestart.Size = new System.Drawing.Size(100, 30);
        this.btnRestart.TabIndex = 4;
        this.btnRestart.Text = "Restart";
        this.btnRestart.UseVisualStyleBackColor = true;
        this.btnRestart.Click += new System.EventHandler(this.btnRestart_Click);

        // btnForceStop
        this.btnForceStop.BackColor = System.Drawing.Color.OrangeRed;
        this.btnForceStop.ForeColor = System.Drawing.Color.White;
        this.btnForceStop.Location = new System.Drawing.Point(580, 25);
        this.btnForceStop.Name = "btnForceStop";
        this.btnForceStop.Size = new System.Drawing.Size(100, 30);
        this.btnForceStop.TabIndex = 5;
        this.btnForceStop.Text = "Force Stop";
        this.btnForceStop.UseVisualStyleBackColor = false;
        this.btnForceStop.Visible = false;

        // progressBar
        this.progressBar.Location = new System.Drawing.Point(250, 65);
        this.progressBar.Name = "progressBar";
        this.progressBar.Size = new System.Drawing.Size(430, 20);
        this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
        this.progressBar.TabIndex = 6;
        this.progressBar.Visible = false;

        // grpTargetFile
        this.grpTargetFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.grpTargetFile.Controls.Add(this.pnlWarningBanner);
        this.grpTargetFile.Controls.Add(this.txtFilePath);
        this.grpTargetFile.Controls.Add(this.btnBrowseFile);
        this.grpTargetFile.Controls.Add(this.lblFileStatus);
        this.grpTargetFile.Controls.Add(this.btnClearFile);
        this.grpTargetFile.Location = new System.Drawing.Point(12, 12);
        this.grpTargetFile.Name = "grpTargetFile";
        this.grpTargetFile.Size = new System.Drawing.Size(760, 100);
        this.grpTargetFile.TabIndex = 1;
        this.grpTargetFile.TabStop = false;
        this.grpTargetFile.Text = "Target File";

        // txtFilePath
        this.txtFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.txtFilePath.Location = new System.Drawing.Point(15, 30);
        this.txtFilePath.Name = "txtFilePath";
        this.txtFilePath.ReadOnly = true;
        this.txtFilePath.Size = new System.Drawing.Size(545, 23);
        this.txtFilePath.TabIndex = 0;
        this.txtFilePath.Text = "(Select a file)";

        // btnBrowseFile
        this.btnBrowseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
        this.btnBrowseFile.Location = new System.Drawing.Point(570, 28);
        this.btnBrowseFile.Name = "btnBrowseFile";
        this.btnBrowseFile.Size = new System.Drawing.Size(85, 27);
        this.btnBrowseFile.TabIndex = 1;
        this.btnBrowseFile.Text = "Browse...";
        this.btnBrowseFile.UseVisualStyleBackColor = true;
        this.btnBrowseFile.Click += new System.EventHandler(this.btnBrowseFile_Click);

        // lblFileStatus
        this.lblFileStatus.AutoSize = true;
        this.lblFileStatus.Location = new System.Drawing.Point(15, 60);
        this.lblFileStatus.Name = "lblFileStatus";
        this.lblFileStatus.Size = new System.Drawing.Size(90, 15);
        this.lblFileStatus.TabIndex = 2;
        this.lblFileStatus.Text = "Status: Not selected";

        // btnClearFile
        this.btnClearFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
        this.btnClearFile.Location = new System.Drawing.Point(665, 28);
        this.btnClearFile.Name = "btnClearFile";
        this.btnClearFile.Size = new System.Drawing.Size(80, 27);
        this.btnClearFile.TabIndex = 3;
        this.btnClearFile.Text = "Clear";
        this.btnClearFile.UseVisualStyleBackColor = true;
        this.btnClearFile.Click += new System.EventHandler(this.btnClearFile_Click);

        // pnlWarningBanner
        this.pnlWarningBanner.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.pnlWarningBanner.BackColor = System.Drawing.Color.Orange;
        this.pnlWarningBanner.Controls.Add(this.lblWarningMessage);
        this.pnlWarningBanner.Location = new System.Drawing.Point(15, 65);
        this.pnlWarningBanner.Name = "pnlWarningBanner";
        this.pnlWarningBanner.Padding = new System.Windows.Forms.Padding(5);
        this.pnlWarningBanner.Size = new System.Drawing.Size(730, 30);
        this.pnlWarningBanner.TabIndex = 4;
        this.pnlWarningBanner.Visible = false;

        // lblWarningMessage
        this.lblWarningMessage.AutoSize = true;
        this.lblWarningMessage.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.lblWarningMessage.ForeColor = System.Drawing.Color.White;
        this.lblWarningMessage.Location = new System.Drawing.Point(5, 7);
        this.lblWarningMessage.Name = "lblWarningMessage";
        this.lblWarningMessage.Size = new System.Drawing.Size(350, 15);
        this.lblWarningMessage.TabIndex = 0;
        this.lblWarningMessage.Text = "⚠ WARNING: File is closed. Please keep the file open.";

        // lblBackupWarning
        this.lblBackupWarning.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.lblBackupWarning.AutoSize = true;
        this.lblBackupWarning.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.lblBackupWarning.ForeColor = System.Drawing.Color.DarkRed;
        this.lblBackupWarning.Location = new System.Drawing.Point(12, 120);
        this.lblBackupWarning.Name = "lblBackupWarning";
        this.lblBackupWarning.Size = new System.Drawing.Size(600, 15);
        this.lblBackupWarning.TabIndex = 3;
        this.lblBackupWarning.Text = "⚠ IMPORTANT: Please create a manual backup before performing any update operations.";

        // grpLogs
        this.grpLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.grpLogs.Controls.Add(this.tabLogs);
        this.grpLogs.Controls.Add(this.btnClearLogs);
        this.grpLogs.Controls.Add(this.btnSaveLogs);
        this.grpLogs.Location = new System.Drawing.Point(12, 254);
        this.grpLogs.Name = "grpLogs";
        this.grpLogs.Size = new System.Drawing.Size(760, 370);
        this.grpLogs.TabIndex = 2;
        this.grpLogs.TabStop = false;
        this.grpLogs.Text = "Log Viewer";

        // tabLogs
        this.tabLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
        this.tabLogs.Controls.Add(this.tabPageServerLog);
        this.tabLogs.Location = new System.Drawing.Point(15, 25);
        this.tabLogs.Name = "tabLogs";
        this.tabLogs.SelectedIndex = 0;
        this.tabLogs.Size = new System.Drawing.Size(730, 330);
        this.tabLogs.TabIndex = 0;

        // tabPageServerLog
        this.tabPageServerLog.Controls.Add(this.txtServerLog);
        this.tabPageServerLog.Location = new System.Drawing.Point(4, 24);
        this.tabPageServerLog.Name = "tabPageServerLog";
        this.tabPageServerLog.Padding = new System.Windows.Forms.Padding(3);
        this.tabPageServerLog.Size = new System.Drawing.Size(722, 302);
        this.tabPageServerLog.TabIndex = 0;
        this.tabPageServerLog.Text = "Server Log";
        this.tabPageServerLog.UseVisualStyleBackColor = true;

        // txtServerLog
        this.txtServerLog.BackColor = System.Drawing.Color.Black;
        this.txtServerLog.Dock = System.Windows.Forms.DockStyle.Fill;
        this.txtServerLog.Font = new System.Drawing.Font("Consolas", 9F);
        this.txtServerLog.ForeColor = System.Drawing.Color.LightGreen;
        this.txtServerLog.Location = new System.Drawing.Point(3, 3);
        this.txtServerLog.Multiline = true;
        this.txtServerLog.Name = "txtServerLog";
        this.txtServerLog.ReadOnly = true;
        this.txtServerLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
        this.txtServerLog.Size = new System.Drawing.Size(716, 296);
        this.txtServerLog.TabIndex = 0;
        this.txtServerLog.WordWrap = false;

        // btnClearLogs
        this.btnClearLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
        this.btnClearLogs.Location = new System.Drawing.Point(15, 365);
        this.btnClearLogs.Name = "btnClearLogs";
        this.btnClearLogs.Size = new System.Drawing.Size(100, 30);
        this.btnClearLogs.TabIndex = 1;
        this.btnClearLogs.Text = "Clear";
        this.btnClearLogs.UseVisualStyleBackColor = true;
        this.btnClearLogs.Click += new System.EventHandler(this.btnClearLogs_Click);

        // btnSaveLogs
        this.btnSaveLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
        this.btnSaveLogs.Location = new System.Drawing.Point(125, 365);
        this.btnSaveLogs.Name = "btnSaveLogs";
        this.btnSaveLogs.Size = new System.Drawing.Size(100, 30);
        this.btnSaveLogs.TabIndex = 2;
        this.btnSaveLogs.Text = "Save...";
        this.btnSaveLogs.UseVisualStyleBackColor = true;
        this.btnSaveLogs.Click += new System.EventHandler(this.btnSaveLogs_Click);

        // MainForm
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(784, 636);
        this.Controls.Add(this.lblBackupWarning);
        this.Controls.Add(this.grpLogs);
        this.Controls.Add(this.grpServerControl);
        this.Controls.Add(this.grpTargetFile);
        this.MinimumSize = new System.Drawing.Size(800, 600);
        this.Name = "MainForm";
        this.Text = "VBA MCP Server Manager";
        this.grpServerControl.ResumeLayout(false);
        this.grpServerControl.PerformLayout();
        this.grpTargetFile.ResumeLayout(false);
        this.grpTargetFile.PerformLayout();
        this.grpLogs.ResumeLayout(false);
        this.tabLogs.ResumeLayout(false);
        this.tabPageServerLog.ResumeLayout(false);
        this.tabPageServerLog.PerformLayout();
        this.ResumeLayout(false);
    }

    #endregion

    private System.Windows.Forms.GroupBox grpServerControl;
    private System.Windows.Forms.Label lblStatus;
    private System.Windows.Forms.Label lblProcessId;
    private System.Windows.Forms.Button btnStart;
    private System.Windows.Forms.Button btnStop;
    private System.Windows.Forms.Button btnRestart;
    private System.Windows.Forms.Button btnForceStop;
    private System.Windows.Forms.ProgressBar progressBar;
    private System.Windows.Forms.GroupBox grpTargetFile;
    private System.Windows.Forms.TextBox txtFilePath;
    private System.Windows.Forms.Button btnBrowseFile;
    private System.Windows.Forms.Label lblFileStatus;
    private System.Windows.Forms.Button btnClearFile;
    private System.Windows.Forms.Panel pnlWarningBanner;
    private System.Windows.Forms.Label lblWarningMessage;
    private System.Windows.Forms.Label lblBackupWarning;
    private System.Windows.Forms.GroupBox grpLogs;
    private System.Windows.Forms.TabControl tabLogs;
    private System.Windows.Forms.TabPage tabPageServerLog;
    private System.Windows.Forms.TextBox txtServerLog;
    private System.Windows.Forms.Button btnClearLogs;
    private System.Windows.Forms.Button btnSaveLogs;
}
