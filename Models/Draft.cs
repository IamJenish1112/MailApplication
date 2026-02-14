using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class Draft
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("subject")]
    public string Subject { get; set; } = string.Empty;

    [BsonElement("body")]
    public string Body { get; set; } = string.Empty;

    [BsonElement("isHtml")]
    public bool IsHtml { get; set; } = true;

    [BsonElement("source")]
    public string Source { get; set; } = "Application";

    [BsonElement("outlookEntryId")]
    public string? OutlookEntryId { get; set; }

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    [BsonElement("updatedAt")]
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
