using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class TemplateImage
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("templateId")]
    public string? TemplateId { get; set; }

    [BsonElement("fileName")]
    public string FileName { get; set; } = string.Empty;

    [BsonElement("mimeType")]
    public string MimeType { get; set; } = "image/png";

    [BsonElement("base64Data")]
    public string? Base64Data { get; set; }

    [BsonElement("filePath")]
    public string? FilePath { get; set; }

    /// <summary>
    /// "base64" | "disk" | "gridfs"
    /// </summary>
    [BsonElement("storageType")]
    public string StorageType { get; set; } = "base64";

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}
