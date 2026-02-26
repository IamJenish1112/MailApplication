namespace MailApplication.Models;

/// <summary>
/// Defines how accounts are rotated during bulk email sending.
/// </summary>
public enum AccountRotationStrategy
{
    /// <summary>
    /// Use a single selected account for all emails.
    /// </summary>
    SingleAccount,

    /// <summary>
    /// Round Robin: Rotate through accounts one email at a time.
    /// Email 1 → Account 1, Email 2 → Account 2, Email 3 → Account 3, Email 4 → Account 1, ...
    /// </summary>
    RoundRobin,

    /// <summary>
    /// Batch Rotation: Send a full batch from one account, then switch to the next.
    /// Batch 1 (50 emails) → Account 1, Batch 2 (50 emails) → Account 2, ...
    /// </summary>
    BatchRotation
}
