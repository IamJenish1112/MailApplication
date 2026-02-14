using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class IndustryEditorForm : Form
{
    private readonly MongoDbService _dbService;
    private readonly Industry? _existingIndustry;
    private TextBox txtName;
    private TextBox txtDescription;
    private Button btnSave;
    private Button btnCancel;

    public IndustryEditorForm(MongoDbService dbService, Industry? existingIndustry)
    {
        _dbService = dbService;
        _existingIndustry = existingIndustry;
        InitializeComponent();
        SetupUI();
        LoadIndustry();
    }

    private void SetupUI()
    {
        this.Text = _existingIndustry == null ? "Add Industry" : "Edit Industry";
        this.Size = new Size(600, 350);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;

        var lblName = new Label
        {
            Text = "Industry Name:",
            Location = new Point(30, 30),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtName = new TextBox
        {
            Location = new Point(30, 60),
            Size = new Size(520, 30),
            Font = new Font("Segoe UI", 10)
        };

        var lblDescription = new Label
        {
            Text = "Description:",
            Location = new Point(30, 110),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtDescription = new TextBox
        {
            Location = new Point(30, 140),
            Size = new Size(520, 80),
            Multiline = true,
            Font = new Font("Segoe UI", 10)
        };

        btnSave = new Button
        {
            Text = "Save",
            Location = new Point(350, 250),
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
            Location = new Point(460, 250),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        this.Controls.AddRange(new Control[] { lblName, txtName, lblDescription, txtDescription, btnSave, btnCancel });
    }

    private void LoadIndustry()
    {
        if (_existingIndustry != null)
        {
            txtName.Text = _existingIndustry.Name;
            txtDescription.Text = _existingIndustry.Description ?? "";
        }
    }

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtName.Text))
        {
            MessageBox.Show("Please enter an industry name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_existingIndustry == null)
        {
            var newIndustry = new Industry
            {
                Name = txtName.Text.Trim(),
                Description = string.IsNullOrWhiteSpace(txtDescription.Text) ? null : txtDescription.Text.Trim(),
                CreatedAt = DateTime.UtcNow
            };

            await _dbService.Industries.InsertOneAsync(newIndustry);
        }
        else
        {
            var filter = Builders<Industry>.Filter.Eq(i => i.Id, _existingIndustry.Id);
            var update = Builders<Industry>.Update
                .Set(i => i.Name, txtName.Text.Trim())
                .Set(i => i.Description, string.IsNullOrWhiteSpace(txtDescription.Text) ? null : txtDescription.Text.Trim());

            await _dbService.Industries.UpdateOneAsync(filter, update);
        }

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
