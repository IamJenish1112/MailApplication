using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class RecipientEditorForm : Form
{
    private readonly MongoDbService _dbService;
    private readonly Recipient? _existingRecipient;

    private TextBox txtEmail;
    private TextBox txtName;
    private CheckedListBox chkIndustries;
    private Button btnSave;
    private Button btnCancel;

    private List<Industry> _industries = new();

    public RecipientEditorForm(MongoDbService dbService, Recipient? existingRecipient)
    {
        _dbService = dbService;
        _existingRecipient = existingRecipient;
        InitializeComponent();
        SetupUI();
        _ = LoadDataAsync();
    }

    private void SetupUI()
    {
        this.Text = _existingRecipient == null ? "Add Recipient" : "Edit Recipient";
        this.Size = new Size(620, 560);
        this.MinimumSize = new Size(500, 450);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.Sizable; // Resizable
        this.MaximizeBox = false;

        // Main scrollable panel
        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Padding = new Padding(25, 20, 25, 20)
        };

        var lblEmail = new Label
        {
            Text = "Email Address:",
            Location = new Point(5, 10),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtEmail = new TextBox
        {
            Location = new Point(5, 38),
            Size = new Size(540, 30),
            Font = new Font("Segoe UI", 10),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var lblName = new Label
        {
            Text = "Name (Optional):",
            Location = new Point(5, 78),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtName = new TextBox
        {
            Location = new Point(5, 106),
            Size = new Size(540, 30),
            Font = new Font("Segoe UI", 10),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var lblIndustries = new Label
        {
            Text = "Industries:",
            Location = new Point(5, 150),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        chkIndustries = new CheckedListBox
        {
            Location = new Point(5, 178),
            Size = new Size(540, 230),
            Font = new Font("Segoe UI", 10),
            CheckOnClick = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        mainPanel.Controls.AddRange(new Control[] { lblEmail, txtEmail, lblName, txtName, lblIndustries, chkIndustries });

        // Bottom button panel (not scrollable)
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(20, 10, 20, 10)
        };

        btnSave = new Button
        {
            Text = "Save",
            Size = new Size(110, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Location = new Point(buttonPanel.Width - 250, 10);
        btnSave.Click += BtnSave_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new Size(110, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Location = new Point(buttonPanel.Width - 130, 10);
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        buttonPanel.Controls.AddRange(new Control[] { btnSave, btnCancel });

        this.Controls.Add(mainPanel);
        this.Controls.Add(buttonPanel);
    }

    private async Task LoadDataAsync()
    {
        await LoadIndustries();
        LoadRecipient();
    }

    private async Task LoadIndustries()
    {
        _industries = await _dbService.Industries.Find(_ => true).ToListAsync();

        chkIndustries.Items.Clear();
        foreach (var industry in _industries)
        {
            chkIndustries.Items.Add(industry.Name);
        }
    }

    private void LoadRecipient()
    {
        if (_existingRecipient != null)
        {
            txtEmail.Text = _existingRecipient.Email;
            txtEmail.ReadOnly = true;
            txtName.Text = _existingRecipient.Name ?? "";

            // Match industries by ObjectId - mark selected industries
            for (int i = 0; i < _industries.Count; i++)
            {
                if (_existingRecipient.Industries.Contains(_industries[i].Id ?? ""))
                {
                    chkIndustries.SetItemChecked(i, true);
                }
            }
        }
    }

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtEmail.Text) || !txtEmail.Text.Contains("@"))
        {
            MessageBox.Show("Please enter a valid email address.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (chkIndustries.CheckedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one industry.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Get selected industry ObjectIds
        var selectedIndustryIds = new List<string>();
        for (int i = 0; i < chkIndustries.Items.Count; i++)
        {
            if (chkIndustries.GetItemChecked(i) && i < _industries.Count)
            {
                var industryId = _industries[i].Id;
                if (!string.IsNullOrEmpty(industryId))
                {
                    selectedIndustryIds.Add(industryId);
                }
            }
        }

        if (_existingRecipient == null)
        {
            var existing = await _dbService.Recipients.Find(r => r.Email == txtEmail.Text.Trim()).FirstOrDefaultAsync();

            if (existing != null)
            {
                // Merge industries (ObjectId-based, deduplicated)
                var mergedIndustries = existing.Industries.Union(selectedIndustryIds).Distinct().ToList();
                var filter = Builders<Recipient>.Filter.Eq(r => r.Id, existing.Id);
                var update = Builders<Recipient>.Update
                    .Set(r => r.Industries, mergedIndustries)
                    .Set(r => r.Name, string.IsNullOrWhiteSpace(txtName.Text) ? existing.Name : txtName.Text.Trim());

                await _dbService.Recipients.UpdateOneAsync(filter, update);
                MessageBox.Show("Email already exists. Industries have been merged.", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                var newRecipient = new Recipient
                {
                    Email = txtEmail.Text.Trim(),
                    Name = string.IsNullOrWhiteSpace(txtName.Text) ? null : txtName.Text.Trim(),
                    Industries = selectedIndustryIds,
                    IsSent = false,
                    CreatedAt = DateTime.UtcNow
                };

                await _dbService.Recipients.InsertOneAsync(newRecipient);
            }
        }
        else
        {
            var filter = Builders<Recipient>.Filter.Eq(r => r.Id, _existingRecipient.Id);
            var update = Builders<Recipient>.Update
                .Set(r => r.Name, string.IsNullOrWhiteSpace(txtName.Text) ? null : txtName.Text.Trim())
                .Set(r => r.Industries, selectedIndustryIds);

            await _dbService.Recipients.UpdateOneAsync(filter, update);
        }

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
