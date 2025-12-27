namespace VbaMcpServer.GUI.Models;

public enum ServerStatus
{
    Stopped,
    Starting,
    Running,
    Stopping
}

public class ServerStatusChangedEventArgs : EventArgs
{
    public ServerStatus Status { get; }
    public int? ProcessId { get; }
    public string? Message { get; }

    public ServerStatusChangedEventArgs(ServerStatus status, int? processId = null, string? message = null)
    {
        Status = status;
        ProcessId = processId;
        Message = message;
    }
}
