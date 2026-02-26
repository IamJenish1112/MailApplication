namespace MailApplication.Models;

/// <summary>
/// Result of a single email send attempt, used for tracking and reporting.
/// </summary>
public class SendResult
{
    public string RecipientEmail { get; set; } = string.Empty;
    public string SenderAccount { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public string? ErrorMessage { get; set; }
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    public int BatchNumber { get; set; }
}

/// <summary>
/// Summary of the entire bulk send operation.
/// </summary>
public class BulkSendSummary
{
    public int TotalEmails { get; set; }
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public int TotalBatches { get; set; }
    public TimeSpan Duration { get; set; }
    public List<SendResult> FailedResults { get; set; } = new();
    public Dictionary<string, int> EmailsPerAccount { get; set; } = new();
}
