using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class AccountsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private ListView listViewAccounts;
    private Button btnAddAccount;
    private Button btnEditAccount;
    private Button btnDeleteAccount;
    private Button btnSetDefault;
    private Button btnRefresh;

    private List<EmailAccount> _accounts = new();

    public AccountsControl(MongoDbService dbService)
    {
        _dbService = dbService;
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
            Text = "Manage email accounts for sending bulk emails. The default account will be used for sending.",
            Location = new Point(0, 40),
            Size = new Size(800, 25),
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

        listViewAccounts.Columns.Add("Account Name", 300);
        listViewAccounts.Columns.Add("Email Address", 400);
        listViewAccounts.Columns.Add("Default", 100);
        listViewAccounts.Columns.Add("Created", 180);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 550),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnAddAccount = CreateButton("Add Account", 0);
        btnEditAccount = CreateButton("Edit", 150);
        btnDeleteAccount = CreateButton("Delete", 270);
        btnSetDefault = CreateButton("Set as Default", 390);
        btnRefresh = CreateButton("Refresh", 550);

        btnAddAccount.Click += BtnAddAccount_Click;
        btnEditAccount.Click += BtnEditAccount_Click;
        btnDeleteAccount.Click += BtnDeleteAccount_Click;
        btnSetDefault.Click += BtnSetDefault_Click;
        btnRefresh.Click += (s, e) => LoadAccounts();

        buttonPanel.Controls.AddRange(new Control[] { btnAddAccount, btnEditAccount, btnDeleteAccount, btnSetDefault, btnRefresh });

        this.Controls.AddRange(new Control[] { titleLabel, infoLabel, listViewAccounts, buttonPanel });
    }

    private Button CreateButton(string text, int left)
    {
        return new Button
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
    }

    private async void LoadAccounts()
    {
        listViewAccounts.Items.Clear();
        _accounts = await _dbService.EmailAccounts.Find(_ => true).ToListAsync();

        foreach (var account in _accounts)
        {
            var item = new ListViewItem(account.AccountName);
            item.SubItems.Add(account.EmailAddress);
            item.SubItems.Add(account.IsDefault ? "Yes" : "No");
            item.SubItems.Add(account.CreatedAt.ToLocalTime().ToString("g"));
            item.Tag = account;
            listViewAccounts.Items.Add(item);
        }
    }

    private void BtnAddAccount_Click(object? sender, EventArgs e)
    {
        var editorForm = new AccountEditorForm(_dbService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
        {
            LoadAccounts();
        }
    }

    private void BtnEditAccount_Click(object? sender, EventArgs e)
    {
        if (listViewAccounts.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an account to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var account = listViewAccounts.SelectedItems[0].Tag as EmailAccount;
        if (account != null)
        {
            var editorForm = new AccountEditorForm(_dbService, account);
            if (editorForm.ShowDialog() == DialogResult.OK)
            {
                LoadAccounts();
            }
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
            await _dbService.EmailAccounts.UpdateManyAsync(
                _ => true,
                Builders<EmailAccount>.Update.Set(a => a.IsDefault, false)
            );

            var filter = Builders<EmailAccount>.Filter.Eq(a => a.Id, account.Id);
            var update = Builders<EmailAccount>.Update.Set(a => a.IsDefault, true);
            await _dbService.EmailAccounts.UpdateOneAsync(filter, update);

            LoadAccounts();
        }
    }
}
