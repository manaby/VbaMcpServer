namespace VbaMcpServer.GUI.Models;

/// <summary>
/// GUI全体の統合状態（11状態のState Machine）
/// </summary>
public enum GuiState
{
    /// <summary>ファイル未選択</summary>
    Idle_NoFile,

    /// <summary>ファイル選択済み（サーバ停止中）</summary>
    Idle_FileSelected,

    /// <summary>起動中: ファイルを開いている（3-13秒）</summary>
    Starting_OpeningFile,

    /// <summary>起動中: ファイルが開くのを待機中（最大10秒）</summary>
    Starting_WaitingForFile,

    /// <summary>起動中: MCPサーバ起動中（1秒）</summary>
    Starting_LaunchingServer,

    /// <summary>実行中: ファイルが開いている正常状態</summary>
    Running_FileOpen,

    /// <summary>実行中: ユーザーがファイルを手動で閉じた（警告状態）</summary>
    Running_FileClosedByUser,

    /// <summary>停止中: サーバプロセスを停止中（0-5秒）</summary>
    Stopping_ServerShutdown,

    /// <summary>停止中: クリーンアップ処理中（瞬時）</summary>
    Stopping_Cleanup,

    /// <summary>エラー: ファイルオープンに失敗</summary>
    Error_FileOpenFailed,

    /// <summary>エラー: サーバプロセスがクラッシュ</summary>
    Error_ServerCrashed,
}

/// <summary>
/// 状態遷移イベント引数
/// </summary>
public class StateChangedEventArgs : EventArgs
{
    public GuiState PreviousState { get; }
    public GuiState NewState { get; }
    public string? Message { get; }

    public StateChangedEventArgs(GuiState previousState, GuiState newState, string? message = null)
    {
        PreviousState = previousState;
        NewState = newState;
        Message = message;
    }
}

/// <summary>
/// 状態遷移を管理するState Machine
/// </summary>
public class GuiStateMachine
{
    private readonly object _lock = new();
    private GuiState _currentState = GuiState.Idle_NoFile;

    /// <summary>現在の状態</summary>
    public GuiState CurrentState
    {
        get { lock (_lock) return _currentState; }
    }

    /// <summary>状態遷移イベント（UIスレッドで発火）</summary>
    public event EventHandler<StateChangedEventArgs>? StateChanged;

    /// <summary>
    /// 指定した状態に遷移可能かチェック
    /// </summary>
    public bool CanTransitionTo(GuiState newState)
    {
        lock (_lock)
        {
            return _validTransitions.ContainsKey(_currentState) &&
                   _validTransitions[_currentState].Contains(newState);
        }
    }

    /// <summary>
    /// 状態遷移を実行（スレッドセーフ）
    /// </summary>
    /// <exception cref="InvalidOperationException">不正な遷移の場合</exception>
    public void TransitionTo(GuiState newState, string? message = null)
    {
        GuiState previousState;

        lock (_lock)
        {
            if (!CanTransitionTo(newState))
            {
                throw new InvalidOperationException(
                    $"Invalid state transition: {_currentState} → {newState}");
            }

            previousState = _currentState;
            _currentState = newState;
        }

        // イベント発火（ロック外で実行）
        StateChanged?.Invoke(this, new StateChangedEventArgs(previousState, newState, message));
    }

    /// <summary>
    /// 状態遷移ルール（許可された遷移のみ定義）
    /// </summary>
    private static readonly Dictionary<GuiState, HashSet<GuiState>> _validTransitions = new()
    {
        [GuiState.Idle_NoFile] = new()
        {
            GuiState.Idle_FileSelected
        },

        [GuiState.Idle_FileSelected] = new()
        {
            GuiState.Idle_NoFile,               // Clear clicked
            GuiState.Starting_OpeningFile       // Start clicked
        },

        [GuiState.Starting_OpeningFile] = new()
        {
            GuiState.Starting_WaitingForFile,   // File open initiated
            GuiState.Error_FileOpenFailed,      // File open failed
            GuiState.Stopping_Cleanup           // Cancel clicked
        },

        [GuiState.Starting_WaitingForFile] = new()
        {
            GuiState.Starting_LaunchingServer,  // File opened successfully
            GuiState.Error_FileOpenFailed,      // Timeout/file didn't open
            GuiState.Stopping_Cleanup           // Cancel clicked
        },

        [GuiState.Starting_LaunchingServer] = new()
        {
            GuiState.Running_FileOpen,          // Server started successfully
            GuiState.Stopping_ServerShutdown    // Cancel clicked
        },

        [GuiState.Running_FileOpen] = new()
        {
            GuiState.Running_FileClosedByUser,  // File closed by user
            GuiState.Stopping_ServerShutdown,   // Stop/Restart clicked
            GuiState.Error_ServerCrashed        // Server crashed
        },

        [GuiState.Running_FileClosedByUser] = new()
        {
            GuiState.Running_FileOpen,          // File reopened
            GuiState.Stopping_ServerShutdown,   // Stop/Restart clicked
            GuiState.Error_ServerCrashed        // Server crashed
        },

        [GuiState.Stopping_ServerShutdown] = new()
        {
            GuiState.Stopping_Cleanup           // Server stopped
        },

        [GuiState.Stopping_Cleanup] = new()
        {
            GuiState.Idle_FileSelected,         // Stop completed
            GuiState.Starting_OpeningFile       // Restart (file needs reopen)
        },

        [GuiState.Error_FileOpenFailed] = new()
        {
            GuiState.Idle_NoFile,               // Clear clicked
            GuiState.Idle_FileSelected,         // Different file selected
            GuiState.Starting_OpeningFile       // Retry
        },

        [GuiState.Error_ServerCrashed] = new()
        {
            GuiState.Idle_NoFile,               // Clear clicked
            GuiState.Idle_FileSelected,         // Different file selected
            GuiState.Starting_OpeningFile       // Restart
        }
    };
}
