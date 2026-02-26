using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class EmailAccount
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("accountName")]
    public string AccountName { get; set; } = string.Empty;

    [BsonElement("emailAddress")]
    public string EmailAddress { get; set; } = string.Empty;

    [BsonElement("smtpAddress")]
    public string SmtpAddress { get; set; } = string.Empty;

    [BsonElement("accountType")]
    public string AccountType { get; set; } = string.Empty;

    [BsonElement("isDefault")]
    public bool IsDefault { get; set; } = false;

    [BsonElement("isEnabled")]
    public bool IsEnabled { get; set; } = true;

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Transient property — tracks how many emails were sent via this account in the current session.
    /// Not persisted to MongoDB.
    /// </summary>
    [BsonIgnore]
    public int SentCount { get; set; } = 0;

    public override string ToString() => $"{AccountName} ({SmtpAddress})";
}
