using MailApplication.Models;

namespace MailApplication.Services;

/// <summary>
/// High-level service coordinating the draft designer workflow.
/// Orchestrates editor bridge, template repository, image storage, and HTML sanitization.
/// </summary>
public class DraftDesignerService
{
    private readonly TemplateRepository _templateRepo;
    private readonly ImageStorageService _imageService;
    private readonly HtmlSanitizerService _sanitizer;

    public DraftDesignerService(MongoDbService dbService)
    {
        _templateRepo = new TemplateRepository(dbService);
        _imageService = new ImageStorageService();
        _sanitizer = new HtmlSanitizerService();
    }

    // ── Template Operations ────────────────────────────────

    public async Task<List<EmailTemplate>> GetAllTemplatesAsync()
        => await _templateRepo.GetAllTemplatesAsync();

    public async Task<EmailTemplate?> GetTemplateByIdAsync(string id)
        => await _templateRepo.GetTemplateByIdAsync(id);

    /// <summary>
    /// Saves a template: cleans HTML, inlines CSS, extracts placeholders.
    /// </summary>
    public async Task<string> SaveTemplateAsync(string? existingId, string name, string subject,
        string category, string rawEditorHtml)
    {
        var cleanedHtml = _sanitizer.CleanHtml(rawEditorHtml);
        var inlinedHtml = _sanitizer.InlineCssForEmail(cleanedHtml);
        var placeholders = _sanitizer.ExtractPlaceholders(cleanedHtml);

        var template = new EmailTemplate
        {
            Id = existingId,
            Name = name,
            Subject = subject,
            Category = category,
            RawEditorHtml = rawEditorHtml,     // Keep raw for re-editing
            HtmlBody = inlinedHtml,             // Email-ready version
            Placeholders = placeholders,
            IsActive = true
        };

        return await _templateRepo.SaveTemplateAsync(template);
    }

    /// <summary>
    /// Converts a template to a Draft for sending.
    /// </summary>
    public Draft ConvertTemplateToDraft(EmailTemplate template, Dictionary<string, string>? placeholderValues = null)
    {
        var htmlBody = template.HtmlBody;

        if (placeholderValues != null && placeholderValues.Count > 0)
        {
            htmlBody = _sanitizer.ReplacePlaceholders(htmlBody, placeholderValues);
        }

        return new Draft
        {
            Subject = placeholderValues != null
                ? _sanitizer.ReplacePlaceholders(template.Subject, placeholderValues)
                : template.Subject,
            Body = htmlBody,
            IsHtml = true,
            Source = "Application"
        };
    }

    public async Task DeleteTemplateAsync(string id)
        => await _templateRepo.DeleteTemplateAsync(id);

    // ── Image Operations ───────────────────────────────────

    /// <summary>
    /// Processes an image file for embedding in the editor.
    /// Returns the base64 data URI for the img src.
    /// </summary>
    public (bool success, string dataUriOrError) ProcessImageForEditor(string filePath)
    {
        var validation = _imageService.ValidateImage(filePath);
        if (!validation.isValid)
            return (false, validation.message);

        var sizeKB = _imageService.GetFileSizeKB(filePath);
        if (sizeKB > 500)
        {
            // Large image warning — still process but caller can warn user
        }

        var dataUri = _imageService.ConvertToBase64DataUri(filePath);
        return (true, dataUri);
    }

    // ── HTML Operations ────────────────────────────────────

    public string CleanHtml(string html) => _sanitizer.CleanHtml(html);

    public string InlineCss(string html) => _sanitizer.InlineCssForEmail(html);

    public string? ValidateHtml(string html) => _sanitizer.ValidateHtml(html);

    public List<string> ExtractPlaceholders(string html) => _sanitizer.ExtractPlaceholders(html);

    /// <summary>
    /// Preview HTML with sample placeholder values.
    /// </summary>
    public string GeneratePreviewHtml(string html)
    {
        var sampleValues = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Company", "Acme Corp" },
            { "OrderId", "ORD-12345" },
            { "Email", "john@example.com" },
            { "Date", DateTime.Now.ToString("MMMM dd, yyyy") },
            { "UnsubscribeLink", "#" }
        };

        var previewHtml = _sanitizer.ReplacePlaceholders(html, sampleValues);
        return _sanitizer.InlineCssForEmail(previewHtml);
    }
}
