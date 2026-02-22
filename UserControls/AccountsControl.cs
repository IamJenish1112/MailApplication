using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class AccountsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;
    private ListView listViewAccounts;
    private Button btnFetchFromOutlook;
    private Button btnSetDefault;
    private Button btnRefresh;
    private Button btnDeleteAccount;

    private List<EmailAccount> _accounts = new();

    public AccountsControl(MongoDbService dbService, OutlookService outlookService)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        InitializeComponent();
        SetupUI();
        LoadAccounts();
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);

        var titleLabel = new Label
        {
            Text = "Email Accounts",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        var infoLabel = new Label
        {
            Text = "Manage email accounts for sending bulk emails. Fetch accounts from Outlook or set a default sending account.",
            Location = new Point(0, 40),
            Size = new Size(900, 25),
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        listViewAccounts = new ListView
        {
            Location = new Point(0, 80),
            Size = new Size(980, 450),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewAccounts.Columns.Add("Account Display Name", 250);
        listViewAccounts.Columns.Add("SMTP Address", 300);
        listViewAccounts.Columns.Add("Account Type", 150);
        listViewAccounts.Columns.Add("Default", 80);
        listViewAccounts.Columns.Add("Created", 180);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 550),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnFetchFromOutlook = new Button
        {
            Text = "⟳ Fetch from Outlook",
            Location = new Point(0, 5),
            Size = new Size(190, 40),
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Enabled = _outlookService.IsAvailable
        };
        btnFetchFromOutlook.FlatAppearance.BorderSize = 0;

        btnSetDefault = CreateButton("Set as Default", 200);
        btnDeleteAccount = CreateButton("Delete", 350);
        btnRefresh = CreateButton("Refresh", 480);

        btnFetchFromOutlook.Click += BtnFetchFromOutlook_Click;
        btnSetDefault.Click += BtnSetDefault_Click;
        btnDeleteAccount.Click += BtnDeleteAccount_Click;
        btnRefresh.Click += (s, e) => LoadAccounts();

        buttonPanel.Controls.AddRange(new Control[] { btnFetchFromOutlook, btnSetDefault, btnDeleteAccount, btnRefresh });

        this.Controls.AddRange(new Control[] { titleLabel, infoLabel, listViewAccounts, buttonPanel });
    }

    private Button CreateButton(string text, int left)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(left, 5),
            Size = new Size(140, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    private async void LoadAccounts()
    {
        listViewAccounts.Items.Clear();
        _accounts = await _dbService.EmailAccounts.Find(_ => true).ToListAsync();

        foreach (var account in _accounts)
        {
            var item = new ListViewItem(account.AccountName);
            item.SubItems.Add(account.SmtpAddress);
            item.SubItems.Add(account.AccountType);
            item.SubItems.Add(account.IsDefault ? "✔ Yes" : "No");
            item.SubItems.Add(account.CreatedAt.ToLocalTime().ToString("g"));
            item.Tag = account;

            if (account.IsDefault)
            {
                item.BackColor = Color.FromArgb(232, 245, 233);
            }

            listViewAccounts.Items.Add(item);
        }
    }

    private async void BtnFetchFromOutlook_Click(object? sender, EventArgs e)
    {
        if (!_outlookService.IsAvailable)
        {
            MessageBox.Show("Outlook is not available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            btnFetchFromOutlook.Enabled = false;
            btnFetchFromOutlook.Text = "Fetching...";

            var outlookAccounts = _outlookService.GetOutlookAccounts();

            if (outlookAccounts.Count == 0)
            {
                MessageBox.Show("No Outlook accounts found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int newCount = 0;
            int updatedCount = 0;

            foreach (var account in outlookAccounts)
            {
                // Check if account already exists by SMTP address
                var existing = await _dbService.EmailAccounts
                    .Find(a => a.SmtpAddress == account.SmtpAddress)
                    .FirstOrDefaultAsync();

                if (existing == null)
                {
                    var newAccount = new EmailAccount
                    {
                        AccountName = account.AccountName,
                        EmailAddress = account.EmailAddress,
                        SmtpAddress = account.SmtpAddress,
                        AccountType = account.AccountType,
                        IsDefault = account.IsDefault && newCount == 0,
                        CreatedAt = DateTime.UtcNow
                    };
                    await _dbService.EmailAccounts.InsertOneAsync(newAccount);
                    newCount++;
                }
                else
                {
                    // Update existing account info
                    var filter = Builders<EmailAccount>.Filter.Eq(a => a.Id, existing.Id);
                    var update = Builders<EmailAccount>.Update
                        .Set(a => a.AccountName, account.AccountName)
                        .Set(a => a.AccountType, account.AccountType);
                    await _dbService.EmailAccounts.UpdateOneAsync(filter, update);
                    updatedCount++;
                }
            }

            MessageBox.Show(
                $"Outlook accounts fetched successfully!\n\nNew accounts: {newCount}\nUpdated accounts: {updatedCount}",
                "Fetch Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            LoadAccounts();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error fetching accounts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnFetchFromOutlook.Enabled = _outlookService.IsAvailable;
            btnFetchFromOutlook.Text = "⟳ Fetch from Outlook";
        }
    }

    private async void BtnSetDefault_Click(object? sender, EventArgs e)
    {
        if (listViewAccounts.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an account to set as default.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var account = listViewAccounts.SelectedItems[0].Tag as EmailAccount;
        if (account != null && !string.IsNullOrEmpty(account.Id))
        {
            // Clear all defaults
            await _dbService.EmailAccounts.UpdateManyAsync(
                _ => true,
                Builders<EmailAccount>.Update.Set(a => a.IsDefault, false)
            );

            // Set selected as default
            var filter = Builders<EmailAccount>.Filter.Eq(a => a.Id, account.Id);
            var update = Builders<EmailAccount>.Update.Set(a => a.IsDefault, true);
            await _dbService.EmailAccounts.UpdateOneAsync(filter, update);

            LoadAccounts();
        }
    }

    private async void BtnDeleteAccount_Click(object? sender, EventArgs e)
    {
        if (listViewAccounts.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an account to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var result = MessageBox.Show("Are you sure you want to delete this account?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            var account = listViewAccounts.SelectedItems[0].Tag as EmailAccount;
            if (account != null && !string.IsNullOrEmpty(account.Id))
            {
                await _dbService.EmailAccounts.DeleteOneAsync(a => a.Id == account.Id);
                LoadAccounts();
            }
        }
    }
}
