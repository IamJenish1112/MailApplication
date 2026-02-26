namespace MailApplication.Services;

/// <summary>
/// Handles image storage for email templates.
/// Supports base64 embedding (for email), disk storage (for large files).
/// </summary>
public class ImageStorageService
{
    private readonly string _imageDirectory;

    public ImageStorageService()
    {
        _imageDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MailApplication", "Images");

        if (!Directory.Exists(_imageDirectory))
            Directory.CreateDirectory(_imageDirectory);
    }

    /// <summary>
    /// Converts an image file to a base64 data URI string.
    /// </summary>
    public string ConvertToBase64DataUri(string filePath)
    {
        var bytes = File.ReadAllBytes(filePath);
        var mimeType = GetMimeType(filePath);
        var base64 = Convert.ToBase64String(bytes);
        return $"data:{mimeType};base64,{base64}";
    }

    /// <summary>
    /// Converts raw bytes to a base64 data URI.
    /// </summary>
    public string ConvertToBase64DataUri(byte[] bytes, string mimeType)
    {
        var base64 = Convert.ToBase64String(bytes);
        return $"data:{mimeType};base64,{base64}";
    }

    /// <summary>
    /// Saves image to disk and returns the file path.
    /// </summary>
    public string SaveToDisk(string sourceFilePath)
    {
        var fileName = $"{Guid.NewGuid()}{Path.GetExtension(sourceFilePath)}";
        var destPath = Path.Combine(_imageDirectory, fileName);
        File.Copy(sourceFilePath, destPath, true);
        return destPath;
    }

    /// <summary>
    /// Saves image bytes to disk and returns the file path.
    /// </summary>
    public string SaveToDisk(byte[] bytes, string extension)
    {
        var fileName = $"{Guid.NewGuid()}.{extension.TrimStart('.')}";
        var destPath = Path.Combine(_imageDirectory, fileName);
        File.WriteAllBytes(destPath, bytes);
        return destPath;
    }

    /// <summary>
    /// Gets file size in kilobytes.
    /// </summary>
    public long GetFileSizeKB(string filePath)
    {
        var info = new FileInfo(filePath);
        return info.Length / 1024;
    }

    /// <summary>
    /// Validates that the file is a supported image format and within size limits.
    /// </summary>
    public (bool isValid, string message) ValidateImage(string filePath, int maxSizeKB = 2048)
    {
        var validExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
        var ext = Path.GetExtension(filePath).ToLowerInvariant();

        if (!validExtensions.Contains(ext))
            return (false, $"Unsupported image format: {ext}. Use JPG, PNG, GIF, BMP, or WebP.");

        var sizeKB = GetFileSizeKB(filePath);
        if (sizeKB > maxSizeKB)
            return (false, $"Image too large ({sizeKB:N0} KB). Maximum allowed: {maxSizeKB:N0} KB.");

        return (true, "Valid");
    }

    public static string GetMimeType(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".png" => "image/png",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".webp" => "image/webp",
            _ => "image/png"
        };
    }
}
