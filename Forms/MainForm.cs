using MailApplication.Services;
using MailApplication.UserControls;

namespace MailApplication.Forms;

public partial class MainForm : Form
{
    private Panel sidebarPanel;
    private Panel contentPanel;
    private Button btnDrafts;
    private Button btnSendMail;
    private Button btnInbox;
    private Button btnOutbox;
    private Button btnRecipients;
    private Button btnAccounts;
    private Button btnSettings;
    private Label lblTitle;

    private MongoDbService _dbService;
    private OutlookService _outlookService;
    private EmailSendingService _emailService;

    private UserControl? _currentControl;
    private Button? _activeButton;

    public MainForm()
    {
        InitializeComponent();
        InitializeServices();
        SetupUI();
        LoadDraftsTab();
    }

    private void InitializeServices()
    {
        _dbService = new MongoDbService();
        _outlookService = new OutlookService();
        _emailService = new EmailSendingService(_dbService, _outlookService);

        if (!_outlookService.IsAvailable)
        {
            MessageBox.Show(
                "Microsoft Outlook is not available on this system.\n\n" +
                "The application will run in limited mode:\n" +
                "- Outlook integration features will be disabled\n" +
                "- You can still create drafts and manage recipients\n" +
                "- Email sending will use application drafts only",
                "Outlook Not Available",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }

    private void SetupUI()
    {
        this.Text = "Bulk Mail Sender - Admin Panel";
        this.Size = new Size(1400, 900);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.FromArgb(248, 249, 250);
        this.MinimumSize = new Size(1200, 700);

        sidebarPanel = new Panel
        {
            Dock = DockStyle.Left,
            Width = 250,
            BackColor = Color.White,
            Padding = new Padding(0)
        };

        lblTitle = new Label
        {
            Text = "BULK MAIL SENDER",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(52, 58, 64),
            AutoSize = false,
            Height = 80,
            Dock = DockStyle.Top,
            TextAlign = ContentAlignment.MiddleCenter,
            BackColor = Color.FromArgb(233, 236, 239)
        };

        sidebarPanel.Controls.Add(lblTitle);

        int buttonTop = 90;
        btnDrafts = CreateMenuButton("Drafts", buttonTop);
        btnSendMail = CreateMenuButton("Send Mail", buttonTop + 60);
        btnInbox = CreateMenuButton("Inbox", buttonTop + 120);
        btnOutbox = CreateMenuButton("Outbox", buttonTop + 180);
        btnRecipients = CreateMenuButton("Recipient List", buttonTop + 240);
        btnAccounts = CreateMenuButton("Accounts", buttonTop + 300);
        btnSettings = CreateMenuButton("Settings", buttonTop + 360);

        btnDrafts.Click += (s, e) => LoadDraftsTab();
        btnSendMail.Click += (s, e) => LoadSendMailTab();
        btnInbox.Click += (s, e) => LoadInboxTab();
        btnOutbox.Click += (s, e) => LoadOutboxTab();
        btnRecipients.Click += (s, e) => LoadRecipientsTab();
        btnAccounts.Click += (s, e) => LoadAccountsTab();
        btnSettings.Click += (s, e) => LoadSettingsTab();

        sidebarPanel.Controls.AddRange(new Control[] { btnDrafts, btnSendMail, btnInbox, btnOutbox, btnRecipients, btnAccounts, btnSettings });

        contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(20),
            AutoScroll = true
        };

        this.Controls.Add(contentPanel);
        this.Controls.Add(sidebarPanel);
    }

    private Button CreateMenuButton(string text, int top)
    {
        var btn = new Button
        {
            Text = text,
            Font = new Font("Segoe UI", 11, FontStyle.Regular),
            ForeColor = Color.FromArgb(73, 80, 87),
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(20, 0, 0, 0),
            Size = new Size(250, 50),
            Location = new Point(0, top),
            Cursor = Cursors.Hand
        };

        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(233, 236, 239);
        btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(222, 226, 230);

        btn.MouseEnter += (s, e) =>
        {
            if (btn != _activeButton)
                btn.BackColor = Color.FromArgb(233, 236, 239);
        };
        btn.MouseLeave += (s, e) =>
        {
            if (btn != _activeButton)
                btn.BackColor = Color.White;
        };

        return btn;
    }

    private void SetActiveButton(Button activeButton)
    {
        foreach (Control control in sidebarPanel.Controls)
        {
            if (control is Button btn)
            {
                btn.BackColor = Color.White;
                btn.ForeColor = Color.FromArgb(73, 80, 87);
                btn.Font = new Font("Segoe UI", 11, FontStyle.Regular);
            }
        }

        activeButton.BackColor = Color.FromArgb(13, 110, 253);
        activeButton.ForeColor = Color.White;
        activeButton.Font = new Font("Segoe UI", 11, FontStyle.Bold);
        _activeButton = activeButton;
    }

    private void LoadUserControl(UserControl control, Button activeButton)
    {
        contentPanel.Controls.Clear();
        _currentControl = control;
        control.Dock = DockStyle.Fill;
        contentPanel.Controls.Add(control);
        SetActiveButton(activeButton);
    }

    private void LoadDraftsTab()
    {
        var draftsControl = new DraftsControl(_dbService, _outlookService);
        LoadUserControl(draftsControl, btnDrafts);
    }

    private void LoadSendMailTab()
    {
        var sendMailControl = new SendMailControl(_dbService, _outlookService, _emailService);
        LoadUserControl(sendMailControl, btnSendMail);
    }

    private void LoadInboxTab()
    {
        var inboxControl = new InboxControl(_outlookService);
        LoadUserControl(inboxControl, btnInbox);
    }

    private void LoadOutboxTab()
    {
        var outboxControl = new OutboxControl(_outlookService);
        LoadUserControl(outboxControl, btnOutbox);
    }

    private void LoadRecipientsTab()
    {
        var recipientsControl = new RecipientsControl(_dbService);
        LoadUserControl(recipientsControl, btnRecipients);
    }

    private void LoadAccountsTab()
    {
        var accountsControl = new AccountsControl(_dbService, _outlookService);
        LoadUserControl(accountsControl, btnAccounts);
    }

    private void LoadSettingsTab()
    {
        var settingsControl = new SettingsControl(_dbService);
        LoadUserControl(settingsControl, btnSettings);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _outlookService?.Dispose();
        base.OnFormClosing(e);
    }
}
