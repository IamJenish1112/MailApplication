using MailApplication.Models;
using MongoDB.Driver;

namespace MailApplication.Services;

/// <summary>
/// Repository for EmailTemplate CRUD operations against MongoDB.
/// </summary>
public class TemplateRepository
{
    private readonly IMongoCollection<EmailTemplate> _templates;
    private readonly IMongoCollection<TemplateImage> _images;

    public TemplateRepository(MongoDbService dbService)
    {
        _templates = dbService.EmailTemplates;
        _images = dbService.TemplateImages;
    }

    // ── Templates ──────────────────────────────────────────

    public async Task<List<EmailTemplate>> GetAllTemplatesAsync()
    {
        return await _templates
            .Find(_ => true)
            .SortByDescending(t => t.UpdatedAt)
            .ToListAsync();
    }

    public async Task<List<EmailTemplate>> GetTemplatesByCategoryAsync(string category)
    {
        return await _templates
            .Find(t => t.Category == category)
            .SortByDescending(t => t.UpdatedAt)
            .ToListAsync();
    }

    public async Task<EmailTemplate?> GetTemplateByIdAsync(string id)
    {
        return await _templates.Find(t => t.Id == id).FirstOrDefaultAsync();
    }

    public async Task<string> SaveTemplateAsync(EmailTemplate template)
    {
        if (string.IsNullOrEmpty(template.Id))
        {
            template.CreatedAt = DateTime.UtcNow;
            template.UpdatedAt = DateTime.UtcNow;
            await _templates.InsertOneAsync(template);
            return template.Id!;
        }
        else
        {
            template.UpdatedAt = DateTime.UtcNow;
            var filter = Builders<EmailTemplate>.Filter.Eq(t => t.Id, template.Id);
            await _templates.ReplaceOneAsync(filter, template);
            return template.Id;
        }
    }

    public async Task DeleteTemplateAsync(string id)
    {
        await _templates.DeleteOneAsync(t => t.Id == id);
        // Also delete associated images
        await _images.DeleteManyAsync(i => i.TemplateId == id);
    }

    // ── Images ─────────────────────────────────────────────

    public async Task<string> SaveImageAsync(TemplateImage image)
    {
        await _images.InsertOneAsync(image);
        return image.Id!;
    }

    public async Task<List<TemplateImage>> GetImagesForTemplateAsync(string templateId)
    {
        return await _images.Find(i => i.TemplateId == templateId).ToListAsync();
    }

    public async Task DeleteImageAsync(string imageId)
    {
        await _images.DeleteOneAsync(i => i.Id == imageId);
    }
}
