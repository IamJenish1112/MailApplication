using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class EmailTemplate
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("name")]
    public string Name { get; set; } = string.Empty;

    [BsonElement("category")]
    public string Category { get; set; } = "General";

    [BsonElement("subject")]
    public string Subject { get; set; } = string.Empty;

    /// <summary>
    /// The final email-ready HTML with inline CSS (output from PreMailer.Net).
    /// </summary>
    [BsonElement("htmlBody")]
    public string HtmlBody { get; set; } = string.Empty;

    /// <summary>
    /// The raw editor HTML (TinyMCE output before inlining) for re-editing.
    /// </summary>
    [BsonElement("rawEditorHtml")]
    public string RawEditorHtml { get; set; } = string.Empty;

    [BsonElement("placeholders")]
    public List<string> Placeholders { get; set; } = new();

    [BsonElement("isActive")]
    public bool IsActive { get; set; } = true;

    [BsonElement("createdAt")]
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    [BsonElement("updatedAt")]
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
