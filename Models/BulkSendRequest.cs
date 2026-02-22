namespace MailApplication.Models;

/// <summary>
/// Encapsulates all parameters for a bulk email send operation.
/// </summary>
public class BulkSendRequest
{
    /// <summary>
    /// The list of recipients to send emails to.
    /// </summary>
    public List<Recipient> Recipients { get; set; } = new();

    /// <summary>
    /// The draft template to use for the email body and subject.
    /// </summary>
    public Draft Draft { get; set; } = null!;

    /// <summary>
    /// Number of recipients per BCC batch (e.g., 50).
    /// </summary>
    public int BatchSize { get; set; } = 50;

    /// <summary>
    /// Delay in seconds between consecutive batches.
    /// </summary>
    public int DelaySeconds { get; set; } = 60;

    /// <summary>
    /// Account rotation strategy to use.
    /// </summary>
    public AccountRotationStrategy RotationStrategy { get; set; } = AccountRotationStrategy.SingleAccount;

    /// <summary>
    /// The SMTP address of the single account to use (only for SingleAccount strategy).
    /// </summary>
    public string? SingleAccountSmtp { get; set; }

    /// <summary>
    /// List of account SMTP addresses to rotate through (for RoundRobin and BatchRotation strategies).
    /// If empty, all available accounts will be used.
    /// </summary>
    public List<string> RotationAccountSmtps { get; set; } = new();
}
