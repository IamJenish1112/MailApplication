using System.Runtime.InteropServices;

namespace MailApplication.Services;

/// <summary>
/// Bridge service for communication between C# (WinForms) and JavaScript (TinyMCE in WebView2).
/// Exposed as a host object to the WebView2 environment.
/// </summary>
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public class EditorBridgeService
{
    /// <summary>Fires when editor content changes (debounced from JS side).</summary>
    public event Action<string>? OnContentChanged;

    /// <summary>Fires when user requests image insert from JS toolbar.</summary>
    public event Action? OnImageInsertRequested;

    /// <summary>Fires when editor is fully initialized and ready.</summary>
    public event Action? OnEditorReady;

    /// <summary>Fires when an error occurs in the editor.</summary>
    public event Action<string>? OnEditorError;

    private string _lastContent = string.Empty;
    private bool _isEditorReady = false;

    public bool IsEditorReady => _isEditorReady;

    // Called from JavaScript when content changes
    public void NotifyContentChanged(string html)
    {
        _lastContent = html ?? string.Empty;
        OnContentChanged?.Invoke(_lastContent);
    }

    // Called from JavaScript when editor is loaded
    public void NotifyEditorReady()
    {
        _isEditorReady = true;
        OnEditorReady?.Invoke();
    }

    // Called from JavaScript on error
    public void NotifyError(string errorMessage)
    {
        OnEditorError?.Invoke(errorMessage ?? "Unknown editor error");
    }

    // Called from JavaScript when user clicks custom image button
    public void RequestImageUpload()
    {
        OnImageInsertRequested?.Invoke();
    }

    public string GetLastContent() => _lastContent;
}
