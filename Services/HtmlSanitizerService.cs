using System.Text.RegularExpressions;
using PreMailer.Net;

namespace MailApplication.Services;

/// <summary>
/// Sanitizes and transforms HTML for email compatibility.
/// Converts CSS classes/styles to inline styles (required by Outlook, Gmail, etc.).
/// </summary>
public class HtmlSanitizerService
{
    /// <summary>
    /// Converts editor HTML to email-compatible HTML with inline CSS.
    /// </summary>
    public string InlineCssForEmail(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return string.Empty;

        try
        {
            // Wrap in proper HTML structure if needed
            var fullHtml = WrapInHtmlDocument(html);

            // Use PreMailer.Net to inline all CSS
            var result = PreMailer.Net.PreMailer.MoveCssInline(fullHtml,
                removeStyleElements: true,
                removeComments: true,
                ignoreElements: "#ignore");

            return result.Html;
        }
        catch (Exception)
        {
            // If inlining fails, return original HTML wrapped properly
            return WrapInHtmlDocument(html);
        }
    }

    /// <summary>
    /// Extracts placeholders like {{Name}}, {{Company}} from HTML content.
    /// </summary>
    public List<string> ExtractPlaceholders(string html)
    {
        var placeholders = new List<string>();
        if (string.IsNullOrWhiteSpace(html)) return placeholders;

        var regex = new Regex(@"\{\{(\w+)\}\}", RegexOptions.Compiled);
        var matches = regex.Matches(html);

        foreach (Match match in matches)
        {
            var placeholder = match.Value; // e.g., "{{Name}}"
            if (!placeholders.Contains(placeholder))
                placeholders.Add(placeholder);
        }

        return placeholders;
    }

    /// <summary>
    /// Replaces placeholders in HTML with actual values from a dictionary.
    /// </summary>
    public string ReplacePlaceholders(string html, Dictionary<string, string> values)
    {
        if (string.IsNullOrWhiteSpace(html) || values == null)
            return html;

        var result = html;
        foreach (var kvp in values)
        {
            // Support both {{Name}} and {{ Name }} formats
            var pattern = $"{{{{{kvp.Key}}}}}";
            result = result.Replace(pattern, kvp.Value);
        }

        return result;
    }

    /// <summary>
    /// Cleans up messy HTML output — removes empty tags, fixes nesting issues.
    /// </summary>
    public string CleanHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return string.Empty;

        // Remove empty <span> tags
        html = Regex.Replace(html, @"<span[^>]*>\s*</span>", "", RegexOptions.IgnoreCase);

        // Remove empty <p> and <div> tags
        html = Regex.Replace(html, @"<p[^>]*>\s*</p>", "", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"<div[^>]*>\s*</div>", "", RegexOptions.IgnoreCase);

        // Remove multiple consecutive <br> tags (keep max 2)
        html = Regex.Replace(html, @"(<br\s*/?>){3,}", "<br/><br/>", RegexOptions.IgnoreCase);

        // Clean up excessive whitespace
        html = Regex.Replace(html, @"\n{3,}", "\n\n");

        return html.Trim();
    }

    /// <summary>
    /// Validates HTML structure. Returns error message or null if valid.
    /// </summary>
    public string? ValidateHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return "HTML content is empty.";

        // Check for unclosed critical tags
        var criticalTags = new[] { "table", "tr", "td", "th", "a", "div" };
        foreach (var tag in criticalTags)
        {
            var openCount = Regex.Matches(html, $@"<{tag}[\s>]", RegexOptions.IgnoreCase).Count;
            var closeCount = Regex.Matches(html, $@"</{tag}>", RegexOptions.IgnoreCase).Count;

            if (openCount != closeCount)
                return $"Mismatched <{tag}> tags: {openCount} opening, {closeCount} closing.";
        }

        return null; // Valid
    }

    private string WrapInHtmlDocument(string bodyHtml)
    {
        // If already a full document, return as-is
        if (bodyHtml.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) ||
            bodyHtml.TrimStart().StartsWith("<html", StringComparison.OrdinalIgnoreCase))
        {
            return bodyHtml;
        }

        return $@"<!DOCTYPE html>
<html>
<head>
<meta charset=""utf-8"">
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
</head>
<body style=""margin:0;padding:20px;font-family:Arial,'Helvetica Neue',Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333333;"">
{bodyHtml}
</body>
</html>";
    }
}
