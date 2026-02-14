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
    private RichTextBox txtBody;
    private CheckBox chkIsHtml;
    private Button btnSave;
    private Button btnCancel;

    public DraftEditorForm(MongoDbService dbService, OutlookService outlookService, Draft? existingDraft)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        _existingDraft = existingDraft;
        InitializeComponent();
        SetupUI();
        LoadDraft();
    }

    private void SetupUI()
    {
        this.Text = _existingDraft == null ? "Create New Draft" : "Edit Draft";
        this.Size = new Size(900, 700);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblSource = new Label
        {
            Text = "Source:",
            Location = new Point(30, 30),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        rbApp = new RadioButton
        {
            Text = "Application",
            Location = new Point(140, 30),
            Size = new Size(120, 25),
            Checked = true,
            Font = new Font("Segoe UI", 10)
        };

        rbOutlook = new RadioButton
        {
            Text = "Outlook",
            Location = new Point(270, 30),
            Size = new Size(120, 25),
            Font = new Font("Segoe UI", 10)
        };

        if (!_outlookService.IsAvailable)
        {
            rbOutlook.Enabled = false;
        }

        var lblSubject = new Label
        {
            Text = "Subject:",
            Location = new Point(30, 80),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtSubject = new TextBox
        {
            Location = new Point(30, 110),
            Size = new Size(820, 30),
            Font = new Font("Segoe UI", 10)
        };

        var lblBody = new Label
        {
            Text = "Body:",
            Location = new Point(30, 160),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtBody = new RichTextBox
        {
            Location = new Point(30, 190),
            Size = new Size(820, 380),
            Font = new Font("Segoe UI", 10)
        };

        chkIsHtml = new CheckBox
        {
            Text = "HTML Format",
            Location = new Point(30, 590),
            Size = new Size(150, 25),
            Checked = true,
            Font = new Font("Segoe UI", 10)
        };

        btnSave = new Button
        {
            Text = "Save",
            Location = new Point(650, 620),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSave.Click += BtnSave_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(760, 620),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        this.Controls.AddRange(new Control[] { lblSource, rbApp, rbOutlook, lblSubject, txtSubject, lblBody, txtBody, chkIsHtml, btnSave, btnCancel });
    }

    private void LoadDraft()
    {
        if (_existingDraft != null)
        {
            txtSubject.Text = _existingDraft.Subject;
            txtBody.Text = _existingDraft.Body;
            chkIsHtml.Checked = _existingDraft.IsHtml;
            rbApp.Checked = _existingDraft.Source == "Application";
            rbOutlook.Checked = _existingDraft.Source == "Outlook";
            rbOutlook.Enabled = false;
            rbApp.Enabled = false;
        }
    }

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSubject.Text))
        {
            MessageBox.Show("Please enter a subject.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            if (rbOutlook.Checked)
            {
                _outlookService.SaveDraftToOutlook(txtSubject.Text, txtBody.Text, chkIsHtml.Checked);
                MessageBox.Show("Draft saved to Outlook successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (_existingDraft == null)
            {
                var newDraft = new Draft
                {
                    Subject = txtSubject.Text,
                    Body = txtBody.Text,
                    IsHtml = chkIsHtml.Checked,
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
                    .Set(d => d.Body, txtBody.Text)
                    .Set(d => d.IsHtml, chkIsHtml.Checked)
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
