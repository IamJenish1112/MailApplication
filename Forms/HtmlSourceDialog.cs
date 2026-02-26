namespace MailApplication.Forms;

/// <summary>
/// Dialog for viewing and editing raw HTML source code.
/// </summary>
public class HtmlSourceDialog : Form
{
    private TextBox txtSource;
    private Button btnApply;
    private Button btnCancel;
    private Button btnFormat;

    public string HtmlContent { get; private set; } = string.Empty;

    public HtmlSourceDialog(string currentHtml)
    {
        HtmlContent = currentHtml;
        SetupUI();
    }

    private void SetupUI()
    {
        this.Text = "HTML Source Editor";
        this.Size = new Size(900, 650);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.FromArgb(30, 30, 30);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MinimumSize = new Size(600, 400);

        var headerLabel = new Label
        {
            Text = "  HTML SOURCE CODE",
            Dock = DockStyle.Top,
            Height = 40,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(200, 200, 200),
            BackColor = Color.FromArgb(40, 40, 40),
            TextAlign = ContentAlignment.MiddleLeft
        };

        txtSource = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font("Cascadia Code", 11),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.FromArgb(220, 220, 170),
            Text = HtmlContent,
            AcceptsTab = true,
            AcceptsReturn = true
        };

        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 55,
            BackColor = Color.FromArgb(40, 40, 40),
            Padding = new Padding(10, 8, 10, 8)
        };

        btnFormat = new Button
        {
            Text = "Format HTML",
            Size = new Size(120, 38),
            Location = new Point(10, 8),
            BackColor = Color.FromArgb(108, 117, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnFormat.FlatAppearance.BorderSize = 0;
        btnFormat.Click += (s, e) => FormatHtml();

        btnApply = new Button
        {
            Text = "Apply",
            Size = new Size(100, 38),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnApply.FlatAppearance.BorderSize = 0;
        btnApply.Click += (s, e) =>
        {
            HtmlContent = txtSource.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        };

        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new Size(100, 38),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) =>
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        };

        // Position buttons
        btnCancel.Location = new Point(buttonPanel.Width - 120, 8);
        btnApply.Location = new Point(buttonPanel.Width - 230, 8);

        buttonPanel.Controls.AddRange(new Control[] { btnFormat, btnApply, btnCancel });

        // Add controls in dock order
        this.Controls.Add(txtSource);
        this.Controls.Add(headerLabel);
        this.Controls.Add(buttonPanel);
    }

    private void FormatHtml()
    {
        try
        {
            var html = txtSource.Text;
            // Simple indentation formatting
            html = html.Replace("><", ">\n<");
            var lines = html.Split('\n');
            var indent = 0;
            var formatted = new System.Text.StringBuilder();
            var selfClosingTags = new[] { "br", "hr", "img", "input", "meta", "link" };

            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                // Decrease indent for closing tags
                if (line.StartsWith("</"))
                    indent = Math.Max(0, indent - 1);

                formatted.AppendLine(new string(' ', indent * 2) + line);

                // Increase indent for opening tags (not self-closing)
                if (line.StartsWith("<") && !line.StartsWith("</") && !line.StartsWith("<!") &&
                    !line.EndsWith("/>") && !selfClosingTags.Any(t => line.StartsWith($"<{t}", StringComparison.OrdinalIgnoreCase)))
                {
                    indent++;
                }
            }

            txtSource.Text = formatted.ToString().TrimEnd();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Format error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
