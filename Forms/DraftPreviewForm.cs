using MailApplication.Models;

namespace MailApplication.Forms;

public partial class DraftPreviewForm : Form
{
    public DraftPreviewForm(Draft draft)
    {
        InitializeComponent();
        SetupUI(draft);
    }

    private void SetupUI(Draft draft)
    {
        this.Text = "Draft Preview";
        this.Size = new Size(800, 600);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;

        var lblSubject = new Label
        {
            Text = "Subject:",
            Location = new Point(20, 20),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        var txtSubject = new TextBox
        {
            Text = draft.Subject,
            Location = new Point(130, 20),
            Size = new Size(640, 30),
            ReadOnly = true,
            Font = new Font("Segoe UI", 10)
        };

        var lblBody = new Label
        {
            Text = "Body:",
            Location = new Point(20, 70),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        var webBrowser = new WebBrowser
        {
            Location = new Point(20, 100),
            Size = new Size(750, 400),
            DocumentText = draft.IsHtml ? draft.Body : $"<pre>{draft.Body}</pre>"
        };

        var btnClose = new Button
        {
            Text = "Close",
            Location = new Point(670, 520),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnClose.Click += (s, e) => this.Close();

        this.Controls.AddRange(new Control[] { lblSubject, txtSubject, lblBody, webBrowser, btnClose });
    }
}
