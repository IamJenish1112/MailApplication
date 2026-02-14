using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class Recipient
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("email")]
    public string Email { get; set; } = string.Empty;

    [BsonElement("name")]
    public string? Name { get; set; }

    [BsonElement("industries")]
    public List<string> Industries { get; set; } = new();

    [BsonElement("isSent")]
    public bool IsSent { get; set; } = false;

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    [BsonElement("lastSentAt")]
    public DateTime? LastSentAt { get; set; }
}
