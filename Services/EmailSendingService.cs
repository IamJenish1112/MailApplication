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

                var uniqueWhiteline = GenerateUniqueWhiteline();
                var bodyWithWhiteline = draft.IsHtml
                    ? $"{draft.Body}<span style='color:white;font-size:1px;'>{uniqueWhiteline}</span>"
                    : $"{draft.Body}\n{uniqueWhiteline}";

                try
                {
                    _outlookService.SendEmail(bccList, draft.Subject, bodyWithWhiteline, draft.IsHtml);

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

    private string GenerateUniqueWhiteline()
    {
        return Guid.NewGuid().ToString("N");
    }
}

public class BatchProgressEventArgs : EventArgs
{
    public int CurrentBatch { get; set; }
    public int TotalBatches { get; set; }
    public int EmailsSent { get; set; }
    public int TotalEmails { get; set; }
}
