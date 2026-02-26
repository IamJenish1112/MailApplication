using MailApplication.Models;
using MongoDB.Driver;

namespace MailApplication.Services;

public class EmailSendingService
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;
    private bool _isSending = false;

    public event EventHandler<BatchProgressEventArgs>? BatchProgress;
    public event EventHandler<string>? StatusUpdate;

    public EmailSendingService(MongoDbService dbService, OutlookService outlookService)
    {
        _dbService = dbService;
        _outlookService = outlookService;
    }

    public bool IsSending => _isSending;

    public async Task SendBulkEmailsAsync(
        List<Recipient> recipients,
        Draft draft,
        int batchSize,
        int delaySeconds,
        string? senderSmtpAddress = null,
        CancellationToken cancellationToken = default)
    {
        _isSending = true;

        try
        {
            var unsentRecipients = recipients.Where(r => !r.IsSent).ToList();
            var totalBatches = (int)Math.Ceiling((double)unsentRecipients.Count / batchSize);
            var currentBatch = 0;

            StatusUpdate?.Invoke(this, $"Starting bulk email send. Total recipients: {unsentRecipients.Count}, Batches: {totalBatches}");

            for (int i = 0; i < unsentRecipients.Count; i += batchSize)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    StatusUpdate?.Invoke(this, "Sending cancelled by user.");
                    break;
                }

                currentBatch++;
                var batch = unsentRecipients.Skip(i).Take(batchSize).ToList();
                var bccList = string.Join(";", batch.Select(r => r.Email));

                // Generate invisible unique identifier per batch
                var uniqueId = GenerateInvisibleUniqueIdentifier();
                var bodyWithUniqueId = InjectInvisibleIdentifier(draft.Body, uniqueId, draft.IsHtml);

                try
                {
                    // Send using BCC only (no TO field), with optional sender account
                    _outlookService.SendEmail(bccList, draft.Subject, bodyWithUniqueId, draft.IsHtml, senderSmtpAddress);

                    foreach (var recipient in batch)
                    {
                        recipient.IsSent = true;
                        recipient.LastSentAt = DateTime.UtcNow;

                        var filter = Builders<Recipient>.Filter.Eq(r => r.Id, recipient.Id);
                        var update = Builders<Recipient>.Update
                            .Set(r => r.IsSent, true)
                            .Set(r => r.LastSentAt, DateTime.UtcNow);

                        await _dbService.Recipients.UpdateOneAsync(filter, update);
                    }

                    BatchProgress?.Invoke(this, new BatchProgressEventArgs
                    {
                        CurrentBatch = currentBatch,
                        TotalBatches = totalBatches,
                        EmailsSent = batch.Count,
                        TotalEmails = unsentRecipients.Count
                    });

                    StatusUpdate?.Invoke(this, $"Batch {currentBatch}/{totalBatches} sent successfully. {batch.Count} emails sent.");

                    if (i + batchSize < unsentRecipients.Count)
                    {
                        StatusUpdate?.Invoke(this, $"Waiting {delaySeconds} seconds before next batch...");
                        await Task.Delay(delaySeconds * 1000, cancellationToken);
                    }
                }
                catch (Exception ex)
                {
                    StatusUpdate?.Invoke(this, $"Error sending batch {currentBatch}: {ex.Message}");
                }
            }

            StatusUpdate?.Invoke(this, "Bulk email sending completed!");
        }
        finally
        {
            _isSending = false;
        }
    }

    /// <summary>
    /// Generates an invisible unique identifier using zero-width Unicode characters
    /// combined with an HTML comment for maximum hash uniqueness.
    /// </summary>
    private string GenerateInvisibleUniqueIdentifier()
    {
        var guid = Guid.NewGuid().ToString("N");
        // Convert GUID chars to zero-width character sequences
        var zeroWidthChars = new[] { "\u200B", "\u200C", "\u200D", "\uFEFF" }; // ZWSP, ZWNJ, ZWJ, BOM
        var encoded = string.Join("", guid.Select(c =>
            zeroWidthChars[Math.Abs(c.GetHashCode()) % zeroWidthChars.Length]));
        return encoded;
    }

    /// <summary>
    /// Injects an invisible unique identifier into the email body.
    /// Uses HTML comment + zero-width characters in a hidden div.
    /// Invisible to recipient but modifies email body hash.
    /// </summary>
    private string InjectInvisibleIdentifier(string body, string uniqueId, bool isHtml)
    {
        var commentTag = $"<!--{Guid.NewGuid():N}-->";

        if (isHtml)
        {
            // Use a combination of: HTML comment + hidden div with zero-width chars
            var hiddenBlock = $"{commentTag}<div style=\"display:none!important;font-size:0;line-height:0;max-height:0;overflow:hidden;mso-hide:all;\">{uniqueId}</div>";

            // Insert before </body> if exists, otherwise append
            if (body.Contains("</body>", StringComparison.OrdinalIgnoreCase))
            {
                var insertIndex = body.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                return body.Insert(insertIndex, hiddenBlock);
            }
            else
            {
                return body + hiddenBlock;
            }
        }
        else
        {
            // For plain text, use zero-width characters only (invisible)
            return body + "\n" + uniqueId;
        }
    }
}

public class BatchProgressEventArgs : EventArgs
{
    public int CurrentBatch { get; set; }
    public int TotalBatches { get; set; }
    public int EmailsSent { get; set; }
    public int TotalEmails { get; set; }
}
