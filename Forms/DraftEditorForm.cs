using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class DraftEditorForm : Form
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;
    private readonly Draft? _existingDraft;

    private RadioButton rbApp;
    private RadioButton rbOutlook;
    private TextBox txtSubject;
    private WebBrowser webEditor;
    private Button btnSave;
    private Button btnCancel;
    private Button btnDesignInOutlook;

    // Toolbar buttons
    private ToolStrip toolStrip;

    public DraftEditorForm(MongoDbService dbService, OutlookService outlookService, Draft? existingDraft)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        _existingDraft = existingDraft;
        InitializeComponent();
        SetupUI();
    }

    private void SetupUI()
    {
        this.Text = _existingDraft == null ? "Create New Draft" : "Edit Draft";
        this.Size = new Size(1000, 780);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MinimumSize = new Size(800, 600);
        this.MaximizeBox = true;
        this.MinimizeBox = false;

        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            AutoScroll = true
        };

        // Source selection
        var sourcePanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 45,
            Padding = new Padding(0, 5, 0, 5)
        };

        var lblSource = new Label
        {
            Text = "Source:",
            Location = new Point(0, 8),
            Size = new Size(70, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        rbApp = new RadioButton
        {
            Text = "Application",
            Location = new Point(80, 8),
            Size = new Size(120, 25),
            Checked = true,
            Font = new Font("Segoe UI", 10)
        };

        rbOutlook = new RadioButton
        {
            Text = "Outlook",
            Location = new Point(210, 8),
            Size = new Size(120, 25),
            Font = new Font("Segoe UI", 10),
            Enabled = _outlookService.IsAvailable
        };

        btnDesignInOutlook = new Button
        {
            Text = "✉ Design in Outlook",
            Location = new Point(360, 3),
            Size = new Size(180, 35),
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Enabled = _outlookService.IsAvailable
        };
        btnDesignInOutlook.FlatAppearance.BorderSize = 0;
        btnDesignInOutlook.Click += BtnDesignInOutlook_Click;

        sourcePanel.Controls.AddRange(new Control[] { lblSource, rbApp, rbOutlook, btnDesignInOutlook });

        // Subject
        var subjectPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 65,
            Padding = new Padding(0, 5, 0, 5)
        };

        var lblSubject = new Label
        {
            Text = "Subject:",
            Location = new Point(0, 8),
            Size = new Size(70, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtSubject = new TextBox
        {
            Location = new Point(0, 33),
            Dock = DockStyle.Bottom,
            Height = 30,
            Font = new Font("Segoe UI", 10)
        };

        subjectPanel.Controls.AddRange(new Control[] { lblSubject, txtSubject });

        // Rich Text Editor Toolbar
        toolStrip = CreateFormattingToolbar();

        // WebBrowser-based rich editor
        var editorPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 5, 0, 0)
        };

        webEditor = new WebBrowser
        {
            Dock = DockStyle.Fill,
            IsWebBrowserContextMenuEnabled = true,
            AllowWebBrowserDrop = true
        };

        webEditor.DocumentCompleted += WebEditor_DocumentCompleted;
        editorPanel.Controls.Add(webEditor);

        // Bottom buttons
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            Padding = new Padding(0, 10, 0, 0)
        };

        btnSave = new Button
        {
            Text = "Save",
            Size = new Size(120, 42),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Location = new Point(buttonPanel.Width - 260, 10);
        btnSave.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnSave.Click += BtnSave_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new Size(120, 42),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Location = new Point(buttonPanel.Width - 130, 10);
        btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        buttonPanel.Controls.AddRange(new Control[] { btnSave, btnCancel });

        // Order matters for docking: Bottom first, then Top items, then Fill
        mainPanel.Controls.Add(editorPanel);       // Fill
        mainPanel.Controls.Add(toolStrip);          // Top (toolbar)
        mainPanel.Controls.Add(subjectPanel);       // Top
        mainPanel.Controls.Add(sourcePanel);        // Top
        mainPanel.Controls.Add(buttonPanel);        // Bottom

        this.Controls.Add(mainPanel);

        // Initialize the HTML editor
        InitializeHtmlEditor();
    }

    private ToolStrip CreateFormattingToolbar()
    {
        var strip = new ToolStrip
        {
            Dock = DockStyle.Top,
            BackColor = Color.FromArgb(248, 249, 250),
            GripStyle = ToolStripGripStyle.Hidden,
            Padding = new Padding(5, 3, 5, 3),
            RenderMode = ToolStripRenderMode.System,
            Height = 38
        };

        // Bold
        var btnBold = new ToolStripButton("B") { Font = new Font("Segoe UI", 10, FontStyle.Bold), ToolTipText = "Bold (Ctrl+B)" };
        btnBold.Click += (s, e) => ExecCommand("bold");

        // Italic
        var btnItalic = new ToolStripButton("I") { Font = new Font("Segoe UI", 10, FontStyle.Italic), ToolTipText = "Italic (Ctrl+I)" };
        btnItalic.Click += (s, e) => ExecCommand("italic");

        // Underline
        var btnUnderline = new ToolStripButton("U") { Font = new Font("Segoe UI", 10, FontStyle.Underline), ToolTipText = "Underline (Ctrl+U)" };
        btnUnderline.Click += (s, e) => ExecCommand("underline");

        strip.Items.Add(btnBold);
        strip.Items.Add(btnItalic);
        strip.Items.Add(btnUnderline);
        strip.Items.Add(new ToolStripSeparator());

        // Font family
        var cmbFont = new ToolStripComboBox("cmbFont")
        {
            Size = new Size(130, 25),
            DropDownStyle = ComboBoxStyle.DropDownList,
            ToolTipText = "Font Family"
        };
        cmbFont.Items.AddRange(new object[] { "Arial", "Calibri", "Cambria", "Courier New", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana" });
        cmbFont.SelectedIndex = 0;
        cmbFont.SelectedIndexChanged += (s, e) => ExecCommand("fontName", cmbFont.SelectedItem?.ToString() ?? "Arial");

        // Font size
        var cmbFontSize = new ToolStripComboBox("cmbFontSize")
        {
            Size = new Size(55, 25),
            DropDownStyle = ComboBoxStyle.DropDownList,
            ToolTipText = "Font Size"
        };
        cmbFontSize.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6", "7" });
        cmbFontSize.SelectedIndex = 2; // Default size 3
        cmbFontSize.SelectedIndexChanged += (s, e) => ExecCommand("fontSize", cmbFontSize.SelectedItem?.ToString() ?? "3");

        strip.Items.Add(cmbFont);
        strip.Items.Add(cmbFontSize);
        strip.Items.Add(new ToolStripSeparator());

        // Text color
        var btnTextColor = new ToolStripButton("A") { Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.Red, ToolTipText = "Text Color" };
        btnTextColor.Click += (s, e) =>
        {
            using var colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                var hexColor = $"#{colorDialog.Color.R:X2}{colorDialog.Color.G:X2}{colorDialog.Color.B:X2}";
                ExecCommand("foreColor", hexColor);
                btnTextColor.ForeColor = colorDialog.Color;
            }
        };

        // Background color
        var btnBackColor = new ToolStripButton("⬛") { ToolTipText = "Highlight Color" };
        btnBackColor.Click += (s, e) =>
        {
            using var colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                var hexColor = $"#{colorDialog.Color.R:X2}{colorDialog.Color.G:X2}{colorDialog.Color.B:X2}";
                ExecCommand("backColor", hexColor);
            }
        };

        strip.Items.Add(btnTextColor);
        strip.Items.Add(btnBackColor);
        strip.Items.Add(new ToolStripSeparator());

        // Alignment
        var btnAlignLeft = new ToolStripButton("⫷") { ToolTipText = "Align Left" };
        btnAlignLeft.Click += (s, e) => ExecCommand("justifyLeft");

        var btnAlignCenter = new ToolStripButton("☰") { ToolTipText = "Center" };
        btnAlignCenter.Click += (s, e) => ExecCommand("justifyCenter");

        var btnAlignRight = new ToolStripButton("⫸") { ToolTipText = "Align Right" };
        btnAlignRight.Click += (s, e) => ExecCommand("justifyRight");

        strip.Items.Add(btnAlignLeft);
        strip.Items.Add(btnAlignCenter);
        strip.Items.Add(btnAlignRight);
        strip.Items.Add(new ToolStripSeparator());

        // Lists
        var btnBulletList = new ToolStripButton("• List") { ToolTipText = "Bullet List" };
        btnBulletList.Click += (s, e) => ExecCommand("insertUnorderedList");

        var btnNumberList = new ToolStripButton("1. List") { ToolTipText = "Numbered List" };
        btnNumberList.Click += (s, e) => ExecCommand("insertOrderedList");

        strip.Items.Add(btnBulletList);
        strip.Items.Add(btnNumberList);
        strip.Items.Add(new ToolStripSeparator());

        // Insert Image
        var btnImage = new ToolStripButton("🖼") { ToolTipText = "Insert Image" };
        btnImage.Click += BtnInsertImage_Click;

        // Insert Link
        var btnLink = new ToolStripButton("🔗") { ToolTipText = "Insert Hyperlink" };
        btnLink.Click += BtnInsertLink_Click;

        // Insert Table
        var btnTable = new ToolStripButton("⊞") { ToolTipText = "Insert Table" };
        btnTable.Click += BtnInsertTable_Click;

        strip.Items.Add(btnImage);
        strip.Items.Add(btnLink);
        strip.Items.Add(btnTable);

        return strip;
    }

    private void InitializeHtmlEditor()
    {
        var htmlTemplate = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 14px;
            margin: 10px;
            line-height: 1.6;
            color: #333;
        }
        table { border-collapse: collapse; }
        td, th { border: 1px solid #ccc; padding: 6px 10px; }
    </style>
</head>
<body contenteditable='true'>
</body>
</html>";

        webEditor.DocumentText = htmlTemplate;
    }

    private void WebEditor_DocumentCompleted(object? sender, WebBrowserDocumentCompletedEventArgs e)
    {
        // Load existing draft content after editor is ready
        if (_existingDraft != null)
        {
            LoadDraft();
        }
    }

    private void LoadDraft()
    {
        if (_existingDraft == null) return;

        txtSubject.Text = _existingDraft.Subject;
        rbApp.Checked = _existingDraft.Source == "Application";
        rbOutlook.Checked = _existingDraft.Source == "Outlook";
        rbOutlook.Enabled = false;
        rbApp.Enabled = false;

        if (webEditor.Document?.Body != null)
        {
            if (_existingDraft.IsHtml)
            {
                webEditor.Document.Body.InnerHtml = _existingDraft.Body;
            }
            else
            {
                webEditor.Document.Body.InnerText = _existingDraft.Body;
            }
        }
    }

    private void ExecCommand(string command, string? value = null)
    {
        if (webEditor.Document != null)
        {
            webEditor.Document.ExecCommand(command, false, value);
            webEditor.Focus();
        }
    }

    private string GetHtmlContent()
    {
        return webEditor.Document?.Body?.InnerHtml ?? string.Empty;
    }

    private void BtnInsertImage_Click(object? sender, EventArgs e)
    {
        using var openFileDialog = new OpenFileDialog
        {
            Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All Files|*.*",
            Title = "Select Image"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var imageBytes = File.ReadAllBytes(openFileDialog.FileName);
                var base64 = Convert.ToBase64String(imageBytes);
                var extension = Path.GetExtension(openFileDialog.FileName).TrimStart('.').ToLower();
                var mimeType = extension switch
                {
                    "jpg" or "jpeg" => "image/jpeg",
                    "png" => "image/png",
                    "gif" => "image/gif",
                    "bmp" => "image/bmp",
                    _ => "image/png"
                };

                var imgTag = $"<img src='data:{mimeType};base64,{base64}' style='max-width:100%;height:auto;' />";
                ExecCommand("insertHTML", imgTag);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting image: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void BtnInsertLink_Click(object? sender, EventArgs e)
    {
        using var linkForm = new Form
        {
            Text = "Insert Hyperlink",
            Size = new Size(500, 220),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            BackColor = Color.White
        };

        var lblUrl = new Label { Text = "URL:", Location = new Point(20, 20), Size = new Size(60, 25), Font = new Font("Segoe UI", 10) };
        var txtUrl = new TextBox { Text = "https://", Location = new Point(20, 45), Size = new Size(440, 30), Font = new Font("Segoe UI", 10) };
        var lblText = new Label { Text = "Display Text:", Location = new Point(20, 80), Size = new Size(120, 25), Font = new Font("Segoe UI", 10) };
        var txtLinkText = new TextBox { Location = new Point(20, 105), Size = new Size(440, 30), Font = new Font("Segoe UI", 10) };
        var btnOk = new Button
        {
            Text = "Insert",
            Location = new Point(280, 140),
            Size = new Size(90, 35),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.OK
        };
        var btnLinkCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(380, 140),
            Size = new Size(80, 35),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.Cancel
        };

        linkForm.Controls.AddRange(new Control[] { lblUrl, txtUrl, lblText, txtLinkText, btnOk, btnLinkCancel });
        linkForm.AcceptButton = btnOk;

        if (linkForm.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(txtUrl.Text))
        {
            var displayText = string.IsNullOrWhiteSpace(txtLinkText.Text) ? txtUrl.Text : txtLinkText.Text;
            var linkHtml = $"<a href='{txtUrl.Text}' target='_blank'>{displayText}</a>";
            ExecCommand("insertHTML", linkHtml);
        }
    }

    private void BtnInsertTable_Click(object? sender, EventArgs e)
    {
        using var tableForm = new Form
        {
            Text = "Insert Table",
            Size = new Size(400, 220),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            BackColor = Color.White
        };

        var lblRows = new Label { Text = "Rows:", Location = new Point(20, 20), Size = new Size(80, 25), Font = new Font("Segoe UI", 10) };
        var numRows = new NumericUpDown { Location = new Point(110, 20), Size = new Size(80, 30), Minimum = 1, Maximum = 50, Value = 3, Font = new Font("Segoe UI", 10) };
        var lblCols = new Label { Text = "Columns:", Location = new Point(20, 60), Size = new Size(80, 25), Font = new Font("Segoe UI", 10) };
        var numCols = new NumericUpDown { Location = new Point(110, 60), Size = new Size(80, 30), Minimum = 1, Maximum = 20, Value = 3, Font = new Font("Segoe UI", 10) };
        var btnOk = new Button
        {
            Text = "Insert",
            Location = new Point(190, 130),
            Size = new Size(90, 35),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.OK
        };
        var btnTableCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(290, 130),
            Size = new Size(80, 35),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.Cancel
        };

        tableForm.Controls.AddRange(new Control[] { lblRows, numRows, lblCols, numCols, btnOk, btnTableCancel });
        tableForm.AcceptButton = btnOk;

        if (tableForm.ShowDialog() == DialogResult.OK)
        {
            int rows = (int)numRows.Value;
            int cols = (int)numCols.Value;

            var tableHtml = "<table style='border-collapse:collapse;width:100%;'>";
            for (int r = 0; r < rows; r++)
            {
                tableHtml += "<tr>";
                for (int c = 0; c < cols; c++)
                {
                    var cellTag = r == 0 ? "th" : "td";
                    var bgStyle = r == 0 ? "background-color:#f0f0f0;font-weight:bold;" : "";
                    tableHtml += $"<{cellTag} style='border:1px solid #ccc;padding:8px;{bgStyle}'>&nbsp;</{cellTag}>";
                }
                tableHtml += "</tr>";
            }
            tableHtml += "</table><br/>";

            ExecCommand("insertHTML", tableHtml);
        }
    }

    private void BtnDesignInOutlook_Click(object? sender, EventArgs e)
    {
        try
        {
            var initialSubject = txtSubject.Text;
            var initialHtml = GetHtmlContent();

            MessageBox.Show(
                "Outlook compose window will open.\n\n" +
                "Design your email template there, then CLOSE the window.\n" +
                "The HTML content will be captured automatically.",
                "Design in Outlook",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            var result = _outlookService.OpenOutlookComposer(initialSubject, initialHtml);

            if (result.HasValue)
            {
                txtSubject.Text = result.Value.subject;
                if (webEditor.Document?.Body != null)
                {
                    webEditor.Document.Body.InnerHtml = result.Value.htmlBody;
                }
                MessageBox.Show("Template imported from Outlook successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSubject.Text))
        {
            MessageBox.Show("Please enter a subject.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var htmlBody = GetHtmlContent();

        if (string.IsNullOrWhiteSpace(htmlBody))
        {
            MessageBox.Show("Please enter email body content.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            if (rbOutlook.Checked)
            {
                _outlookService.SaveDraftToOutlook(txtSubject.Text, htmlBody, true);
                MessageBox.Show("Draft saved to Outlook successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (_existingDraft == null)
            {
                var newDraft = new Draft
                {
                    Subject = txtSubject.Text,
                    Body = htmlBody,
                    IsHtml = true, // Always save as HTML
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
                    .Set(d => d.Body, htmlBody)
                    .Set(d => d.IsHtml, true) // Always HTML
                    .Set(d => d.UpdatedAt, DateTime.UtcNow);

                await _dbService.Drafts.UpdateOneAsync(filter, update);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving draft: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
