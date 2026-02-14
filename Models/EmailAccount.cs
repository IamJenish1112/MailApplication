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

    [BsonElement("isDefault")]
    public bool IsDefault { get; set; } = false;

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}
