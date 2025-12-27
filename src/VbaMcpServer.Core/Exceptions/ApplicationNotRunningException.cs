namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when the required Office application is not running
/// </summary>
public class ApplicationNotRunningException : VbaMcpException
{
    public ApplicationNotRunningException(string applicationName)
        : base($"{applicationName} is not running. Please open {applicationName} and try again.",
               "APPLICATION_NOT_RUNNING")
    {
    }
}
