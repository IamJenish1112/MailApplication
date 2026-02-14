using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MailApplication.Models;

public class AppSettings
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string? Id { get; set; }

    [BsonElement("batchSize")]
    public int BatchSize { get; set; } = 50;

    [BsonElement("delayBetweenBatches")]
    public int DelayBetweenBatches { get; set; } = 60;

    [BsonElement("mongoConnectionString")]
    public string MongoConnectionString { get; set; } = "mongodb://localhost:27017";

    [BsonElement("databaseName")]
    public string DatabaseName { get; set; } = "BulkMailSender";

    [BsonElement("updatedAt")]
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
