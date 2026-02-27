using MongoDB.Driver;
using MailApplication.Models;

namespace MailApplication.Services;

public class MongoDbService
{
    private readonly IMongoDatabase _database;

    public MongoDbService(string connectionString = "mongodb://localhost:27017", string databaseName = "BulkMailSender")
    {
        var client = new MongoClient(connectionString);
        _database = client.GetDatabase(databaseName);
    }

    public IMongoCollection<Recipient> Recipients => _database.GetCollection<Recipient>("recipients");
    public IMongoCollection<Draft> Drafts => _database.GetCollection<Draft>("drafts");
    public IMongoCollection<Industry> Industries => _database.GetCollection<Industry>("industries");
    public IMongoCollection<EmailAccount> EmailAccounts => _database.GetCollection<EmailAccount>("emailAccounts");
    public IMongoCollection<AppSettings> Settings => _database.GetCollection<AppSettings>("settings");
    public IMongoCollection<EmailTemplate> EmailTemplates => _database.GetCollection<EmailTemplate>("emailTemplates");
    public IMongoCollection<TemplateImage> TemplateImages => _database.GetCollection<TemplateImage>("templateImages");
    public IMongoCollection<ReplyToAccount> ReplyToAccounts => _database.GetCollection<ReplyToAccount>("replyToAccounts");
}