using MailApplication.Models;
using MailApplication.Services;

namespace MailApplication.UserControls;

public partial class AccountsControl : UserControl
{
    private readonly OutlookAccountService _accountService;
    private ListView listViewAccounts;

    // ── Buttons that are hidden/commented-out per request ──────────────────
    // private Button btnFetchFromOutlook;
    // private Button btnSetDefault;
    // ────────────────────────────────────────────────────────────────────────

    private Button btnRefresh;
    private Label lblStatus;

    private List<EmailAccount> _accounts = new();

    public AccountsControl(OutlookAccountService accountService)
    {
        _accountService = accountService;
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
            Text = "Configured Outlook accounts used for sending bulk emails via round-robin rotation.",
            Location = new Point(0, 40),
            Size = new Size(900, 25),
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        lblStatus = new Label
        {
            Text = "",
            Location = new Point(0, 68),
            Size = new Size(980, 22),
            Font = new Font("Segoe UI", 9, FontStyle.Italic),
            ForeColor = Color.FromArgb(25, 135, 84)
        };

        listViewAccounts = new ListView
        {
            Location = new Point(0, 95),
            Size = new Size(980, 450),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewAccounts.Columns.Add("#", 40);
        listViewAccounts.Columns.Add("Display Name", 240);
        listViewAccounts.Columns.Add("SMTP / Email Address", 300);
        listViewAccounts.Columns.Add("Account Type", 150);
        listViewAccounts.Columns.Add("Sending Order", 150);

        // ── Button panel ───────────────────────────────────────────────────
        var buttonPanel = new Panel
        {
            Location = new Point(0, 560),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        // Fetch from Outlook — hidden (commented out)
        // btnFetchFromOutlook = new Button { ... };

        // Set as Default — hidden (commented out)
        // btnSetDefault = CreateButton("Set as Default", 200);

        btnRefresh = new Button
        {
            Text = "⟳  Refresh",
            Location = new Point(0, 5),
            Size = new Size(140, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnRefresh.FlatAppearance.BorderSize = 0;
        btnRefresh.Click += (s, e) => LoadAccounts();

        buttonPanel.Controls.Add(btnRefresh);

        this.Controls.AddRange(new Control[] { titleLabel, infoLabel, lblStatus, listViewAccounts, buttonPanel });
    }

    private void LoadAccounts()
    {
        listViewAccounts.Items.Clear();

        // Same source that SendMailControl uses — OutlookAccountService cache
        _accounts = _accountService.GetAllAccounts();

        if (_accounts.Count == 0)
        {
            lblStatus.Text = "⚠  No Outlook accounts found. Make sure Outlook is open and configured.";
            lblStatus.ForeColor = Color.FromArgb(220, 53, 69);
        }
        else if (_accounts.Count == 1)
        {
            lblStatus.Text = $"📧  1 account configured — emails will be sent via this account.";
            lblStatus.ForeColor = Color.FromArgb(13, 110, 253);
        }
        else
        {
            lblStatus.Text = $"🔄  {_accounts.Count} accounts configured — round-robin rotation will be used when sending.";
            lblStatus.ForeColor = Color.FromArgb(25, 135, 84);
        }

        int index = 1;
        foreach (var account in _accounts)
        {
            var item = new ListViewItem(index.ToString());
            item.SubItems.Add(account.AccountName);
            item.SubItems.Add(account.SmtpAddress);
            item.SubItems.Add(account.AccountType);
            item.SubItems.Add(index == 1 ? "Primary (1st)" : $"Rotation #{index}");
            item.Tag = account;

            // Alternate row shading
            if (index % 2 == 0)
                item.BackColor = Color.FromArgb(248, 249, 250);

            listViewAccounts.Items.Add(item);
            index++;
        }
    }
}
