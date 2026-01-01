namespace VbaMcpServer.Exceptions;

/// <summary>
/// Exception thrown when a specified report is not found in the Access database
/// </summary>
public class ReportNotFoundException : VbaMcpException
{
    public ReportNotFoundException(string reportName, string? filePath = null)
        : base($"Report not found: {reportName}", "REPORT_NOT_FOUND", filePath)
    {
    }

    public ReportNotFoundException(string reportName, Exception innerException, string? filePath = null)
        : base($"Report not found: {reportName}", "REPORT_NOT_FOUND", innerException, filePath)
    {
    }
}
