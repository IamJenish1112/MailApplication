using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class AccountEditorForm : Form
{
    private readonly MongoDbService _dbService;
    private readonly EmailAccount? _existingAccount;

    private TextBox txtAccountName;
    private TextBox txtEmailAddress;
    private CheckBox chkIsDefault;
    private Button btnSave;
    private Button btnCancel;

    public AccountEditorForm(MongoDbService dbService, EmailAccount? existingAccount)
    {
        _dbService = dbService;
        _existingAccount = existingAccount;
        InitializeComponent();
        SetupUI();
        LoadAccount();
    }

    private void SetupUI()
    {
        this.Text = _existingAccount == null ? "Add Email Account" : "Edit Email Account";
        this.Size = new Size(600, 350);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;

        var lblAccountName = new Label
        {
            Text = "Account Name:",
            Location = new Point(30, 30),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtAccountName = new TextBox
        {
            Location = new Point(30, 60),
            Size = new Size(520, 30),
            Font = new Font("Segoe UI", 10)
        };

        var lblEmailAddress = new Label
        {
            Text = "Email Address:",
            Location = new Point(30, 110),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtEmailAddress = new TextBox
        {
            Location = new Point(30, 140),
            Size = new Size(520, 30),
            Font = new Font("Segoe UI", 10)
        };

        chkIsDefault = new CheckBox
        {
            Text = "Set as default account",
            Location = new Point(30, 190),
            Size = new Size(200, 25),
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

        this.Controls.AddRange(new Control[] { lblAccountName, txtAccountName, lblEmailAddress, txtEmailAddress, chkIsDefault, btnSave, btnCancel });
    }

    private void LoadAccount()
    {
        if (_existingAccount != null)
        {
            txtAccountName.Text = _existingAccount.AccountName;
            txtEmailAddress.Text = _existingAccount.EmailAddress;
            chkIsDefault.Checked = _existingAccount.IsDefault;
        }
    }

    private async void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtAccountName.Text))
        {
            MessageBox.Show("Please enter an account name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtEmailAddress.Text) || !txtEmailAddress.Text.Contains("@"))
        {
            MessageBox.Show("Please enter a valid email address.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_existingAccount == null)
        {
            if (chkIsDefault.Checked)
            {
                await _dbService.EmailAccounts.UpdateManyAsync(
                    _ => true,
                    Builders<EmailAccount>.Update.Set(a => a.IsDefault, false)
                );
            }

            var newAccount = new EmailAccount
            {
                AccountName = txtAccountName.Text.Trim(),
                EmailAddress = txtEmailAddress.Text.Trim(),
                IsDefault = chkIsDefault.Checked,
                CreatedAt = DateTime.UtcNow
            };

            await _dbService.EmailAccounts.InsertOneAsync(newAccount);
        }
        else
        {
            if (chkIsDefault.Checked && !_existingAccount.IsDefault)
            {
                await _dbService.EmailAccounts.UpdateManyAsync(
                    _ => true,
                    Builders<EmailAccount>.Update.Set(a => a.IsDefault, false)
                );
            }

            var filter = Builders<EmailAccount>.Filter.Eq(a => a.Id, _existingAccount.Id);
            var update = Builders<EmailAccount>.Update
                .Set(a => a.AccountName, txtAccountName.Text.Trim())
                .Set(a => a.EmailAddress, txtEmailAddress.Text.Trim())
                .Set(a => a.IsDefault, chkIsDefault.Checked);

            await _dbService.EmailAccounts.UpdateOneAsync(filter, update);
        }

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
