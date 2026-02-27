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
    private Button btnSettings;
    private Label lblTitle;

    private MongoDbService _dbService;
    private OutlookService _outlookService;
    private EmailSendingService _emailService;
    private OutlookAccountService _accountService;
    private BulkMailSenderService _bulkSenderService;

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
        _accountService = new OutlookAccountService();
        _bulkSenderService = new BulkMailSenderService(_dbService, _accountService);

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

        // ── Sidebar ──────────────────────────────────────────────────────────
        sidebarPanel = new Panel
        {
            Dock = DockStyle.Left,
            Width = 240,
            BackColor = Color.FromArgb(24, 28, 36),
            Padding = new Padding(0)
        };

        // App logo / title area
        var logoPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 90,
            BackColor = Color.FromArgb(13, 110, 253)
        };

        lblTitle = new Label
        {
            Text = "BULK MAIL\nSENDER",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };

        logoPanel.Controls.Add(lblTitle);
        sidebarPanel.Controls.Add(logoPanel);

        // Nav buttons (top-to-bottom after header)
        int top = 100;
        btnDrafts   = CreateMenuButton("📋  Drafts",    top);          top += 56;
        btnSendMail = CreateMenuButton("✉  Send Mail",  top);          top += 56;
        btnInbox    = CreateMenuButton("📥  Inbox",      top);          top += 56;
        btnOutbox   = CreateMenuButton("📤  Sent",     top);          top += 56;
        btnSettings = CreateMenuButton("⚙  Settings",   top);

        btnDrafts.Click   += (s, e) => LoadDraftsTab();
        btnSendMail.Click += (s, e) => LoadSendMailTab();
        btnInbox.Click    += (s, e) => LoadInboxTab();
        btnOutbox.Click   += (s, e) => LoadOutboxTab();
        btnSettings.Click += (s, e) => LoadSettingsTab();

        sidebarPanel.Controls.AddRange(new Control[] { btnDrafts, btnSendMail, btnInbox, btnOutbox, btnSettings });

        // ── Content area ─────────────────────────────────────────────────────
        contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(10),
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
            ForeColor = Color.FromArgb(180, 190, 210),
            BackColor = Color.FromArgb(24, 28, 36),
            FlatStyle = FlatStyle.Flat,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(22, 0, 0, 0),
            Size = new Size(240, 52),
            Location = new Point(0, top),
            Cursor = Cursors.Hand
        };

        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(35, 42, 55);
        btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(13, 110, 253);

        btn.MouseEnter += (s, e) =>
        {
            if (btn != _activeButton)
                btn.BackColor = Color.FromArgb(35, 42, 55);
        };
        btn.MouseLeave += (s, e) =>
        {
            if (btn != _activeButton)
                btn.BackColor = Color.FromArgb(24, 28, 36);
        };

        return btn;
    }

    private void SetActiveButton(Button activeButton)
    {
        foreach (Control control in sidebarPanel.Controls)
        {
            if (control is Button btn)
            {
                btn.BackColor = Color.FromArgb(24, 28, 36);
                btn.ForeColor = Color.FromArgb(180, 190, 210);
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
        var sendMailControl = new SendMailControl(_dbService, _outlookService, _emailService, _accountService, _bulkSenderService);
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

    private void LoadSettingsTab()
    {
        var settingsControl = new SettingsControl(_dbService, _accountService);
        LoadUserControl(settingsControl, btnSettings);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _outlookService?.Dispose();
        _accountService?.Dispose();
        base.OnFormClosing(e);
    }
}
