using MailApplication.Models;
using MailApplication.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using MongoDB.Driver;

namespace MailApplication.Forms;

/// <summary>
/// Professional HTML email template editor using WebView2 + TinyMCE.
/// Replaces the legacy WebBrowser-based editor.
/// </summary>
public partial class DraftEditorForm : Form
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;
    private readonly Draft? _existingDraft;
    private readonly DraftDesignerService _designerService;
    private readonly EditorBridgeService _bridgeService;
    private readonly ImageStorageService _imageService;

    // UI Controls
    private TextBox txtSubject = null!;
    private WebView2 webView = null!;
    private Button btnSave = null!;
    private Button btnCancel = null!;
    private Button btnPreview = null!;
    private Button btnHtmlSource = null!;
    private Button btnDesignInOutlook = null!;
    private Panel loadingPanel = null!;
    private Label lblStatus = null!;
    private RadioButton rbApp = null!;
    private RadioButton rbOutlook = null!;

    private bool _isEditorReady = false;
    private string _pendingContent = string.Empty;
    private int _initRetryCount = 0;
    private const int MaxRetries = 3;

    public DraftEditorForm(MongoDbService dbService, OutlookService outlookService, Draft? existingDraft)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        _existingDraft = existingDraft;
        _designerService = new DraftDesignerService(dbService);
        _bridgeService = new EditorBridgeService();
        _imageService = new ImageStorageService();

        InitializeComponent();
        SetupUI();
        _ = InitializeWebView2Async();
    }

    private void SetupUI()
    {
        this.Text = _existingDraft == null ? "✉ Create Email Template" : "✉ Edit Email Template";
        this.Size = new Size(1100, 850);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.FromArgb(248, 249, 250);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MinimumSize = new Size(900, 650);
        this.MaximizeBox = true;
        this.MinimizeBox = false;
        this.Font = new Font("Segoe UI", 10);

        var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

        // ── Header Panel ──────────────────────────────────────
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top, Height = 50, BackColor = Color.White,
            Padding = new Padding(15, 0, 15, 0)
        };
        headerPanel.Paint += (s, e) =>
        {
            using var pen = new Pen(Color.FromArgb(222, 226, 230));
            e.Graphics.DrawLine(pen, 0, headerPanel.Height - 1, headerPanel.Width, headerPanel.Height - 1);
        };

        var headerTitle = new Label
        {
            Text = _existingDraft == null ? "New Email Template" : "Edit Template",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true, Location = new Point(15, 12)
        };
        headerPanel.Controls.Add(headerTitle);

        // ── Source Selection Panel ─────────────────────────────
        var sourcePanel = new Panel { Dock = DockStyle.Top, Height = 45 };
        var lblSource = new Label
        {
            Text = "Source:", Location = new Point(0, 10), AutoSize = true,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        rbApp = new RadioButton
        {
            Text = "Application", Location = new Point(70, 9),
            AutoSize = true, Checked = true, Font = new Font("Segoe UI", 10)
        };
        rbOutlook = new RadioButton
        {
            Text = "Outlook", Location = new Point(190, 9),
            AutoSize = true, Font = new Font("Segoe UI", 10),
            Enabled = _outlookService.IsAvailable
        };
        btnDesignInOutlook = new Button
        {
            Text = "✉ Design in Outlook", Location = new Point(310, 4),
            Size = new Size(175, 34),
            BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand, Enabled = _outlookService.IsAvailable
        };
        btnDesignInOutlook.FlatAppearance.BorderSize = 0;
        btnDesignInOutlook.Click += BtnDesignInOutlook_Click;
        sourcePanel.Controls.AddRange(new Control[] { lblSource, rbApp, rbOutlook, btnDesignInOutlook });

        // ── Subject Panel ─────────────────────────────────────
        var subjectPanel = new Panel { Dock = DockStyle.Top, Height = 60 };
        var lblSubject = new Label
        {
            Text = "Subject:", Location = new Point(0, 5), AutoSize = true,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        txtSubject = new TextBox
        {
            Location = new Point(0, 28), Height = 30,
            Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 11),
            BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle
        };
        subjectPanel.Controls.AddRange(new Control[] { lblSubject, txtSubject });

        // ── Editor Panel (WebView2) ───────────────────────────
        var editorPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 8, 0, 0) };

        webView = new WebView2
        {
            Dock = DockStyle.Fill,
            DefaultBackgroundColor = Color.FromArgb(248, 249, 250)
        };
        webView.CoreWebView2InitializationCompleted += WebView_InitializationCompleted;

        // Loading overlay
        loadingPanel = new Panel
        {
            Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 249, 250), Visible = true
        };
        var loadingLabel = new Label
        {
            Text = "⏳ Loading editor...",
            Font = new Font("Segoe UI", 12), ForeColor = Color.FromArgb(108, 117, 125),
            AutoSize = true
        };
        loadingLabel.Location = new Point(
            (editorPanel.Width - loadingLabel.Width) / 2,
            (editorPanel.Height - loadingLabel.Height) / 2);
        loadingLabel.Anchor = AnchorStyles.None;
        loadingPanel.Controls.Add(loadingLabel);

        editorPanel.Controls.Add(loadingPanel);
        editorPanel.Controls.Add(webView);

        // ── Bottom Bar ────────────────────────────────────────
        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom, Height = 60,
            BackColor = Color.White, Padding = new Padding(10, 10, 10, 10)
        };
        bottomPanel.Paint += (s, e) =>
        {
            using var pen = new Pen(Color.FromArgb(222, 226, 230));
            e.Graphics.DrawLine(pen, 0, 0, bottomPanel.Width, 0);
        };

        // Status label
        lblStatus = new Label
        {
            Text = "Ready", Location = new Point(10, 15),
            AutoSize = true, ForeColor = Color.FromArgb(108, 117, 125),
            Font = new Font("Segoe UI", 9)
        };

        btnHtmlSource = CreateBottomButton("{ } Source", Color.FromArgb(108, 117, 253));
        btnHtmlSource.Click += BtnHtmlSource_Click;

        btnPreview = CreateBottomButton("👁 Preview", Color.FromArgb(23, 162, 184));
        btnPreview.Click += BtnPreview_Click;

        btnSave = CreateBottomButton("💾 Save", Color.FromArgb(40, 167, 69));
        btnSave.Click += BtnSave_Click;

        btnCancel = CreateBottomButton("Cancel", Color.FromArgb(108, 117, 125));
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

        // Position from right
        btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnSave.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnPreview.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnHtmlSource.Anchor = AnchorStyles.Right | AnchorStyles.Top;

        btnCancel.Location = new Point(bottomPanel.Width - 120, 10);
        btnSave.Location = new Point(bottomPanel.Width - 240, 10);
        btnPreview.Location = new Point(bottomPanel.Width - 370, 10);
        btnHtmlSource.Location = new Point(bottomPanel.Width - 500, 10);

        bottomPanel.Controls.AddRange(new Control[] { lblStatus, btnHtmlSource, btnPreview, btnSave, btnCancel });

        // ── Assemble Layout ───────────────────────────────────
        mainPanel.Controls.Add(editorPanel);       // Fill
        mainPanel.Controls.Add(subjectPanel);      // Top
        mainPanel.Controls.Add(sourcePanel);       // Top
        mainPanel.Controls.Add(bottomPanel);       // Bottom

        this.Controls.Add(mainPanel);
        this.Controls.Add(headerPanel);            // Top-most
    }

    private Button CreateBottomButton(string text, Color bgColor)
    {
        var btn = new Button
        {
            Text = text, Size = new Size(120, 38),
            BackColor = bgColor, ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // ══════════════════════════════════════════════════════════
    // WebView2 Initialization
    // ══════════════════════════════════════════════════════════

    private async Task InitializeWebView2Async()
    {
        try
        {
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MailApplication", "WebView2Data");

            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);

            // Add the bridge host object for JS ↔ C# communication
            webView.CoreWebView2.AddHostObjectToScript("bridge", _bridgeService);

            // Wire up bridge events
            _bridgeService.OnEditorReady += () =>
            {
                if (this.InvokeRequired)
                    this.Invoke(OnEditorReady);
                else
                    OnEditorReady();
            };

            _bridgeService.OnImageInsertRequested += () =>
            {
                if (this.InvokeRequired)
                    this.Invoke(HandleImageInsert);
                else
                    HandleImageInsert();
            };

            _bridgeService.OnEditorError += (msg) =>
            {
                if (this.InvokeRequired)
                    this.Invoke(() => UpdateStatus($"⚠ Editor error: {msg}", Color.OrangeRed));
                else
                    UpdateStatus($"⚠ Editor error: {msg}", Color.OrangeRed);
            };

            // Handle WebView2 process failures
            webView.CoreWebView2.ProcessFailed += (s, e) =>
            {
                this.Invoke(() =>
                {
                    UpdateStatus("⚠ Editor crashed. Restarting...", Color.OrangeRed);
                    if (_initRetryCount < MaxRetries)
                    {
                        _initRetryCount++;
                        _ = InitializeWebView2Async();
                    }
                    else
                    {
                        MessageBox.Show(
                            "The editor crashed and could not be restarted.\nPlease close and reopen this form.",
                            "Editor Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            };

            // Navigate to the editor HTML page
            var editorPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "editor", "editor.html");
            if (File.Exists(editorPath))
            {
                webView.CoreWebView2.Navigate(new Uri(editorPath).AbsoluteUri);
            }
            else
            {
                // Fallback: try project directory
                var projectEditorPath = Path.Combine(
                    Directory.GetCurrentDirectory(), "wwwroot", "editor", "editor.html");
                if (File.Exists(projectEditorPath))
                {
                    webView.CoreWebView2.Navigate(new Uri(projectEditorPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show(
                        "Editor HTML file not found.\nExpected at: " + editorPath,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        catch (Exception ex)
        {
            loadingPanel.Visible = false;
            MessageBox.Show(
                $"Failed to initialize WebView2:\n{ex.Message}\n\n" +
                "Please ensure WebView2 Runtime is installed.\n" +
                "Download: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
                "WebView2 Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void WebView_InitializationCompleted(object? sender, CoreWebView2InitializationCompletedEventArgs e)
    {
        if (!e.IsSuccess)
        {
            UpdateStatus("⚠ WebView2 initialization failed", Color.OrangeRed);
            if (e.InitializationException != null)
            {
                MessageBox.Show(
                    $"WebView2 init error:\n{e.InitializationException.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // ══════════════════════════════════════════════════════════
    // Editor Ready & Content Management
    // ══════════════════════════════════════════════════════════

    private void OnEditorReady()
    {
        _isEditorReady = true;
        loadingPanel.Visible = false;
        UpdateStatus("✓ Editor ready", Color.FromArgb(40, 167, 69));

        // Load existing draft content
        if (_existingDraft != null)
        {
            LoadDraftContent();
        }

        // Load pending content if any
        if (!string.IsNullOrEmpty(_pendingContent))
        {
            _ = SetEditorContentAsync(_pendingContent);
            _pendingContent = string.Empty;
        }
    }

    private void LoadDraftContent()
    {
        if (_existingDraft == null) return;

        txtSubject.Text = _existingDraft.Subject;
        rbApp.Checked = _existingDraft.Source == "Application";
        rbOutlook.Checked = _existingDraft.Source == "Outlook";

        if (!string.IsNullOrEmpty(_existingDraft.Body))
        {
            _ = SetEditorContentAsync(_existingDraft.Body);
        }
    }

    /// <summary>Sets HTML content in the TinyMCE editor via JavaScript.</summary>
    private async Task SetEditorContentAsync(string html)
    {
        if (!_isEditorReady)
        {
            _pendingContent = html;
            return;
        }

        try
        {
            // Escape backticks and backslashes for JS template literal
            var escapedHtml = html
                .Replace("\\", "\\\\")
                .Replace("`", "\\`")
                .Replace("$", "\\$");

            await webView.CoreWebView2.ExecuteScriptAsync(
                $"setEditorContent(`{escapedHtml}`)");
        }
        catch (Exception ex)
        {
            UpdateStatus($"⚠ Error setting content: {ex.Message}", Color.OrangeRed);
        }
    }

    /// <summary>Gets HTML content from the TinyMCE editor via JavaScript.</summary>
    private async Task<string> GetEditorContentAsync()
    {
        if (!_isEditorReady)
            return string.Empty;

        try
        {
            var result = await webView.CoreWebView2.ExecuteScriptAsync("getEditorContent()");

            // Result comes back as a JSON string (with quotes), so we need to unescape
            if (result != null && result.StartsWith("\"") && result.EndsWith("\""))
            {
                result = result[1..^1]; // Remove surrounding quotes
                result = System.Text.RegularExpressions.Regex.Unescape(result);
            }

            return result ?? string.Empty;
        }
        catch (Exception ex)
        {
            UpdateStatus($"⚠ Error getting content: {ex.Message}", Color.OrangeRed);
            return string.Empty;
        }
    }

    // ══════════════════════════════════════════════════════════
    // Button Event Handlers
    // ══════════════════════════════════════════════════════════

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSubject.Text))
        {
            MessageBox.Show("Please enter a subject.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var htmlBody = await GetEditorContentAsync();
        if (string.IsNullOrWhiteSpace(htmlBody))
        {
            MessageBox.Show("Please enter email body content.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Validate HTML
        var validationError = _designerService.ValidateHtml(htmlBody);
        if (validationError != null)
        {
            var proceed = MessageBox.Show(
                $"HTML Warning:\n{validationError}\n\nSave anyway?",
                "HTML Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (proceed != DialogResult.Yes) return;
        }

        try
        {
            btnSave.Enabled = false;
            UpdateStatus("Saving...", Color.FromArgb(0, 120, 212));

            // Clean and process the HTML
            var cleanedHtml = _designerService.CleanHtml(htmlBody);

            if (rbOutlook.Checked)
            {
                _outlookService.SaveDraftToOutlook(txtSubject.Text, cleanedHtml, true);
                MessageBox.Show("Draft saved to Outlook.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (_existingDraft == null)
            {
                var newDraft = new Draft
                {
                    Subject = txtSubject.Text,
                    Body = cleanedHtml,
                    IsHtml = true,
                    Source = "Application",
                    CreatedAt = DateTime.UtcNow,
                    UpdatedAt = DateTime.UtcNow
                };
                await _dbService.Drafts.InsertOneAsync(newDraft);
            }
            else
            {
                var filter = Builders<Draft>.Filter.Eq(d => d.Id, _existingDraft.Id);
                var update = Builders<Draft>.Update
                    .Set(d => d.Subject, txtSubject.Text)
                    .Set(d => d.Body, cleanedHtml)
                    .Set(d => d.IsHtml, true)
                    .Set(d => d.UpdatedAt, DateTime.UtcNow);
                await _dbService.Drafts.UpdateOneAsync(filter, update);
            }

            UpdateStatus("✓ Saved successfully", Color.FromArgb(40, 167, 69));
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            UpdateStatus("⚠ Save failed", Color.OrangeRed);
            MessageBox.Show($"Error saving draft:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnSave.Enabled = true;
        }
    }

    private async void BtnPreview_Click(object? sender, EventArgs e)
    {
        var html = await GetEditorContentAsync();
        if (string.IsNullOrWhiteSpace(html))
        {
            MessageBox.Show("No content to preview.", "Preview", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var previewHtml = _designerService.GeneratePreviewHtml(html);
        var previewDraft = new Draft
        {
            Subject = txtSubject.Text,
            Body = previewHtml,
            IsHtml = true
        };

        using var previewForm = new DraftPreviewForm(previewDraft);
        previewForm.ShowDialog();
    }

    private async void BtnHtmlSource_Click(object? sender, EventArgs e)
    {
        var html = await GetEditorContentAsync();
        using var sourceDialog = new HtmlSourceDialog(html);

        if (sourceDialog.ShowDialog() == DialogResult.OK)
        {
            await SetEditorContentAsync(sourceDialog.HtmlContent);
            UpdateStatus("✓ HTML source applied", Color.FromArgb(40, 167, 69));
        }
    }

    private async void HandleImageInsert()
    {
        using var openFileDialog = new OpenFileDialog
        {
            Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.webp|All Files|*.*",
            Title = "Select Image for Email"
        };

        if (openFileDialog.ShowDialog() != DialogResult.OK) return;

        try
        {
            var (success, dataUriOrError) = _designerService.ProcessImageForEditor(openFileDialog.FileName);

            if (!success)
            {
                MessageBox.Show(dataUriOrError, "Image Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Warn about large images
            var sizeKB = _imageService.GetFileSizeKB(openFileDialog.FileName);
            if (sizeKB > 500)
            {
                var proceed = MessageBox.Show(
                    $"This image is {sizeKB:N0} KB. Large images may cause email delivery issues.\n\nInsert anyway?",
                    "Large Image", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (proceed != DialogResult.Yes) return;
            }

            var imgHtml = $"<img src=\"{dataUriOrError}\" style=\"max-width:100%;height:auto;\" alt=\"Email Image\" />";
            var escapedImg = imgHtml.Replace("\\", "\\\\").Replace("`", "\\`").Replace("$", "\\$");

            await webView.CoreWebView2.ExecuteScriptAsync(
                $"insertContentAtCursor(`{escapedImg}`)");

            UpdateStatus($"✓ Image inserted ({sizeKB:N0} KB)", Color.FromArgb(40, 167, 69));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error inserting image:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void BtnDesignInOutlook_Click(object? sender, EventArgs e)
    {
        try
        {
            var initialSubject = txtSubject.Text;
            var initialHtml = await GetEditorContentAsync();

            MessageBox.Show(
                "Outlook compose window will open.\n\nDesign your email there, then CLOSE the window.\nThe HTML will be captured automatically.",
                "Design in Outlook", MessageBoxButtons.OK, MessageBoxIcon.Information);

            var result = _outlookService.OpenOutlookComposer(initialSubject, initialHtml);

            if (result.HasValue)
            {
                txtSubject.Text = result.Value.subject;
                await SetEditorContentAsync(result.Value.htmlBody);
                UpdateStatus("✓ Template imported from Outlook", Color.FromArgb(40, 167, 69));
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Outlook error:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ══════════════════════════════════════════════════════════
    // Helpers
    // ══════════════════════════════════════════════════════════

    private void UpdateStatus(string message, Color color)
    {
        if (lblStatus != null)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = color;
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        try
        {
            webView?.Dispose();
        }
        catch { /* Ignore disposal errors */ }
        base.OnFormClosing(e);
    }
}
