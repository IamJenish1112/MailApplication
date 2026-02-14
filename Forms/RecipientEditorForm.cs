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

    public RecipientEditorForm(MongoDbService dbService, Recipient? existingRecipient)
    {
        _dbService = dbService;
        _existingRecipient = existingRecipient;
        InitializeComponent();
        SetupUI();
        LoadIndustries();
        LoadRecipient();
    }

    private void SetupUI()
    {
        this.Text = _existingRecipient == null ? "Add Recipient" : "Edit Recipient";
        this.Size = new Size(600, 500);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;

        var lblEmail = new Label
        {
            Text = "Email:",
            Location = new Point(30, 30),
            Size = new Size(100, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtEmail = new TextBox
        {
            Location = new Point(30, 60),
            Size = new Size(520, 30),
            Font = new Font("Segoe UI", 10)
        };

        var lblName = new Label
        {
            Text = "Name (Optional):",
            Location = new Point(30, 110),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtName = new TextBox
        {
            Location = new Point(30, 140),
            Size = new Size(520, 30),
            Font = new Font("Segoe UI", 10)
        };

        var lblIndustries = new Label
        {
            Text = "Industries:",
            Location = new Point(30, 190),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        chkIndustries = new CheckedListBox
        {
            Location = new Point(30, 220),
            Size = new Size(520, 180),
            Font = new Font("Segoe UI", 10),
            CheckOnClick = true
        };

        btnSave = new Button
        {
            Text = "Save",
            Location = new Point(350, 420),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnSave.Click += BtnSave_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(460, 420),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        this.Controls.AddRange(new Control[] { lblEmail, txtEmail, lblName, txtName, lblIndustries, chkIndustries, btnSave, btnCancel });
    }

    private async void LoadIndustries()
    {
        var industries = await _dbService.Industries.Find(_ => true).ToListAsync();
        
        foreach (var industry in industries)
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

            for (int i = 0; i < chkIndustries.Items.Count; i++)
            {
                if (_existingRecipient.Industries.Contains(chkIndustries.Items[i].ToString() ?? ""))
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

        var selectedIndustries = chkIndustries.CheckedItems.Cast<string>().ToList();

        if (_existingRecipient == null)
        {
            var existing = await _dbService.Recipients.Find(r => r.Email == txtEmail.Text).FirstOrDefaultAsync();
            
            if (existing != null)
            {
                var updatedIndustries = existing.Industries.Union(selectedIndustries).Distinct().ToList();
                var filter = Builders<Recipient>.Filter.Eq(r => r.Id, existing.Id);
                var update = Builders<Recipient>.Update
                    .Set(r => r.Industries, updatedIndustries)
                    .Set(r => r.Name, string.IsNullOrWhiteSpace(txtName.Text) ? existing.Name : txtName.Text);
                
                await _dbService.Recipients.UpdateOneAsync(filter, update);
                MessageBox.Show("Email already exists. Industries have been merged.", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                var newRecipient = new Recipient
                {
                    Email = txtEmail.Text.Trim(),
                    Name = string.IsNullOrWhiteSpace(txtName.Text) ? null : txtName.Text.Trim(),
                    Industries = selectedIndustries,
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
                .Set(r => r.Industries, selectedIndustries);

            await _dbService.Recipients.UpdateOneAsync(filter, update);
        }

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
