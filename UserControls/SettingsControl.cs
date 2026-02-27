using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class SettingsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private readonly OutlookAccountService _accountService;

    // ── Industry Management tab ──────────────────────────────────────────────
    private ListView listViewIndustries;
    private Button btnAddIndustry;
    private Button btnEditIndustry;
    private Button btnDeleteIndustry;
    private Button btnRefreshIndustries;
    private List<Industry> _industries = new();

    // ── App Settings tab ─────────────────────────────────────────────────────
    private TextBox txtBatchSize;
    private TextBox txtDelay;
    private Button btnSaveSettings;

    // ── Recipients tab ───────────────────────────────────────────────────────
    private ComboBox cmbIndustryFilter;
    private ListView listViewRecipients;
    private Button btnAddRecipient;
    private Button btnEditRecipient;
    private Button btnDeleteRecipient;
    private Button btnImportRecipients;
    private Button btnRefreshRecipients;
    private TextBox txtSearch;
    private List<Recipient> _allRecipients = new();
    private List<Industry> _allIndustries = new();

    // ── Accounts tab ─────────────────────────────────────────────────────────
    private ListView listViewAccounts;
    private Button btnRefreshAccounts;
    private Label lblAccountStatus;
    private List<EmailAccount> _accounts = new();

    public SettingsControl(MongoDbService dbService, OutlookAccountService? accountService = null)
    {
        _dbService = dbService;
        _accountService = accountService ?? new OutlookAccountService();
        InitializeComponent();
        SetupUI();
        LoadAllData();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  UI SETUP
    // ═══════════════════════════════════════════════════════════════════════════

    private void SetupUI()
    {
        this.BackColor = Color.FromArgb(248, 249, 250);
        this.Padding = new Padding(0);

        // ── Header bar ──────────────────────────────────────────────────────
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 70,
            BackColor = Color.White,
            Padding = new Padding(24, 0, 0, 0)
        };

        var titleLabel = new Label
        {
            Text = "Settings",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(24, 16)
        };


        headerPanel.Controls.AddRange(new Control[] { titleLabel});

        // ── Tab control ──────────────────────────────────────────────────────
        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10, FontStyle.Regular),
            Padding = new Point(16, 6)
        };

        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem += TabControl_DrawItem;

        var tabRecipients = new TabPage { Text = "  Recipients  ", BackColor = Color.White };
        var tabAccounts   = new TabPage { Text = "  Accounts  ",   BackColor = Color.White };
        var tabIndustries = new TabPage { Text = "  Industry Management  ", BackColor = Color.White };
        var tabAppSettings = new TabPage { Text = "  App Settings  ", BackColor = Color.White };

        BuildRecipientsTab(tabRecipients);
        BuildAccountsTab(tabAccounts);
        BuildIndustriesTab(tabIndustries);
        BuildAppSettingsTab(tabAppSettings);

        tabControl.TabPages.AddRange(new[] { tabRecipients, tabAccounts, tabIndustries, tabAppSettings });

        this.Controls.Add(tabControl);
        this.Controls.Add(headerPanel);
    }

    // ─── Custom tab header drawing ────────────────────────────────────────────
    private void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
    {
        var tc = (TabControl)sender!;
        var tab = tc.TabPages[e.Index];
        bool selected = e.Index == tc.SelectedIndex;

        var bgColor = selected ? Color.White : Color.FromArgb(240, 242, 245);
        var fgColor = selected ? Color.FromArgb(13, 110, 253) : Color.FromArgb(73, 80, 87);
        var font    = new Font("Segoe UI", 10, selected ? FontStyle.Bold : FontStyle.Regular);

        e.Graphics.FillRectangle(new SolidBrush(bgColor), e.Bounds);

        if (selected)
        {
            // Bottom accent bar
            e.Graphics.FillRectangle(
                new SolidBrush(Color.FromArgb(13, 110, 253)),
                new Rectangle(e.Bounds.Left, e.Bounds.Bottom - 3, e.Bounds.Width, 3));
        }

        var textRect = new RectangleF(e.Bounds.Left, e.Bounds.Top + 4, e.Bounds.Width, e.Bounds.Height - 4);
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        e.Graphics.DrawString(tab.Text, font, new SolidBrush(fgColor), textRect, sf);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  TAB: RECIPIENTS
    // ═══════════════════════════════════════════════════════════════════════════

    private void BuildRecipientsTab(TabPage page)
    {
        page.Padding = new Padding(16);
        page.AutoScroll = true;

        var lblTitle = MakeSectionLabel("Recipient Management", new Point(0, 0));

        var filterPanel = new Panel
        {
            Location = new Point(0, 44),
            Size = new Size(980, 35),
            BackColor = Color.White
        };

        var lblIndustry = new Label
        {
            Text = "Filter by Industry:",
            Location = new Point(0, 6),
            Size = new Size(140, 22),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbIndustryFilter = new ComboBox
        {
            Location = new Point(148, 3),
            Size = new Size(230, 28),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };
        cmbIndustryFilter.SelectedIndexChanged += (s, e) => FilterRecipients();

        var lblSearch = new Label
        {
            Text = "Search:",
            Location = new Point(400, 6),
            Size = new Size(65, 22),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtSearch = new TextBox
        {
            Location = new Point(470, 3),
            Size = new Size(250, 28),
            Font = new Font("Segoe UI", 10)
        };
        txtSearch.TextChanged += (s, e) => FilterRecipients();

        filterPanel.Controls.AddRange(new Control[] { lblIndustry, cmbIndustryFilter, lblSearch, txtSearch });

        listViewRecipients = new ListView
        {
            Location = new Point(0, 94),
            Size = new Size(980, 420),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10),
            BorderStyle = BorderStyle.FixedSingle
        };
        listViewRecipients.Columns.Add("Email", 290);
        listViewRecipients.Columns.Add("Name", 190);
        listViewRecipients.Columns.Add("Industries", 270);
        listViewRecipients.Columns.Add("Status", 100);
        listViewRecipients.Columns.Add("Created", 100);

        var btnPanel = new Panel
        {
            Location = new Point(0, 528),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnAddRecipient      = MakeTabButton("+ Add Recipient", 0);
        btnEditRecipient     = MakeTabButton("✏ Edit", 160);
        btnDeleteRecipient   = MakeTabButton("🗑 Delete", 285, Color.FromArgb(220, 53, 69));
        btnImportRecipients  = MakeTabButton("📂 Import File", 420);
        btnRefreshRecipients = MakeTabButton("⟳ Refresh", 570);

        btnAddRecipient.Click      += BtnAddRecipient_Click;
        btnEditRecipient.Click     += BtnEditRecipient_Click;
        btnDeleteRecipient.Click   += BtnDeleteRecipient_Click;
        btnImportRecipients.Click  += BtnImportRecipients_Click;
        btnRefreshRecipients.Click += async (s, e) => { await LoadIndustries(); await LoadRecipients(); };

        btnPanel.Controls.AddRange(new Control[] { btnAddRecipient, btnEditRecipient, btnDeleteRecipient, btnImportRecipients, btnRefreshRecipients });

        page.Controls.AddRange(new Control[] { lblTitle, filterPanel, listViewRecipients, btnPanel });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  TAB: ACCOUNTS
    // ═══════════════════════════════════════════════════════════════════════════

    private void BuildAccountsTab(TabPage page)
    {
        page.Padding = new Padding(16);

        var lblTitle = MakeSectionLabel("Email Accounts", new Point(0, 0));

        var infoLabel = new Label
        {
            Text = "Configured Outlook accounts used for sending bulk emails via round-robin rotation.",
            Location = new Point(0, 44),
            Size = new Size(900, 22),
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        lblAccountStatus = new Label
        {
            Text = "",
            Location = new Point(0, 70),
            Size = new Size(980, 22),
            Font = new Font("Segoe UI", 9, FontStyle.Italic),
            ForeColor = Color.FromArgb(25, 135, 84)
        };

        listViewAccounts = new ListView
        {
            Location = new Point(0, 100),
            Size = new Size(980, 430),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10),
            BorderStyle = BorderStyle.FixedSingle
        };
        listViewAccounts.Columns.Add("#", 40);
        listViewAccounts.Columns.Add("Display Name", 240);
        listViewAccounts.Columns.Add("SMTP / Email Address", 300);
        listViewAccounts.Columns.Add("Account Type", 150);
        listViewAccounts.Columns.Add("Sending Order", 150);

        btnRefreshAccounts = MakeTabButton("⟳ Refresh", 0);
        btnRefreshAccounts.Location = new Point(0, 545);
        btnRefreshAccounts.Click += (s, e) => LoadAccounts();

        page.Controls.AddRange(new Control[] { lblTitle, infoLabel, lblAccountStatus, listViewAccounts, btnRefreshAccounts });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  TAB: INDUSTRY MANAGEMENT
    // ═══════════════════════════════════════════════════════════════════════════

    private void BuildIndustriesTab(TabPage page)
    {
        page.Padding = new Padding(16);

        var lblTitle = MakeSectionLabel("Industry Management", new Point(0, 0));

        listViewIndustries = new ListView
        {
            Location = new Point(0, 50),
            Size = new Size(980, 370),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10),
            BorderStyle = BorderStyle.FixedSingle
        };
        listViewIndustries.Columns.Add("Industry Name", 500);
        listViewIndustries.Columns.Add("Description", 350);
        listViewIndustries.Columns.Add("Created", 130);

        var btnPanel = new Panel
        {
            Location = new Point(0, 434),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnAddIndustry    = MakeTabButton("+ Add Industry", 0);
        btnEditIndustry   = MakeTabButton("✏ Edit", 160);
        btnDeleteIndustry = MakeTabButton("🗑 Delete", 315, Color.FromArgb(220, 53, 69));
        btnRefreshIndustries = MakeTabButton("⟳ Refresh", 460);

        btnAddIndustry.Click    += BtnAddIndustry_Click;
        btnEditIndustry.Click   += BtnEditIndustry_Click;
        btnDeleteIndustry.Click += BtnDeleteIndustry_Click;
        btnRefreshIndustries.Click += async (s, e) => await LoadIndustriesForTab();

        btnPanel.Controls.AddRange(new Control[] { btnAddIndustry, btnEditIndustry, btnDeleteIndustry, btnRefreshIndustries });

        page.Controls.AddRange(new Control[] { lblTitle, listViewIndustries, btnPanel });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  TAB: APP SETTINGS
    // ═══════════════════════════════════════════════════════════════════════════

    private void BuildAppSettingsTab(TabPage page)
    {
        page.Padding = new Padding(16);

        var lblTitle = MakeSectionLabel("Campaign Settings", new Point(0, 0));

        // Card panel
        var card = new Panel
        {
            Location = new Point(0, 54),
            Size = new Size(560, 220),
            BackColor = Color.White,
            Padding = new Padding(20)
        };
        card.Paint += (s, e) =>
        {
            var pen = new Pen(Color.FromArgb(222, 226, 230), 1);
            e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
        };

        var lblBatchSize = new Label
        {
            Text = "Default Batch Size:",
            Location = new Point(20, 30),
            Size = new Size(180, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtBatchSize = new TextBox
        {
            Location = new Point(210, 28),
            Size = new Size(120, 28),
            Text = "50",
            Font = new Font("Segoe UI", 10)
        };

        var lblBatchHint = new Label
        {
            Text = "Number of emails per batch",
            Location = new Point(340, 31),
            Size = new Size(200, 22),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        var lblDelay = new Label
        {
            Text = "Delay Between Batches:",
            Location = new Point(20, 80),
            Size = new Size(180, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtDelay = new TextBox
        {
            Location = new Point(210, 78),
            Size = new Size(120, 28),
            Text = "60",
            Font = new Font("Segoe UI", 10)
        };

        var lblDelayHint = new Label
        {
            Text = "Seconds to wait between batches",
            Location = new Point(340, 81),
            Size = new Size(200, 22),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        var divider = new Panel
        {
            Location = new Point(20, 120),
            Size = new Size(520, 1),
            BackColor = Color.FromArgb(222, 226, 230)
        };

        btnSaveSettings = new Button
        {
            Text = "💾  Save Settings",
            Location = new Point(20, 140),
            Size = new Size(170, 42),
            BackColor = Color.FromArgb(25, 135, 84),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSaveSettings.FlatAppearance.BorderSize = 0;
        btnSaveSettings.Click += BtnSaveSettings_Click;

        card.Controls.AddRange(new Control[]
        {
            lblBatchSize, txtBatchSize, lblBatchHint,
            lblDelay, txtDelay, lblDelayHint,
            divider, btnSaveSettings
        });

        page.Controls.AddRange(new Control[] { lblTitle, card });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  HELPERS
    // ═══════════════════════════════════════════════════════════════════════════

    private Label MakeSectionLabel(string text, Point location)
    {
        return new Label
        {
            Text = text,
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = location
        };
    }

    private Button MakeTabButton(string text, int left, Color? color = null)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(left, 5),
            Size = new Size(148, 38),
            BackColor = color ?? Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  DATA LOADING
    // ═══════════════════════════════════════════════════════════════════════════

    private async void LoadAllData()
    {
        await LoadIndustries();
        await LoadRecipients();
        LoadAccounts();
        await LoadIndustriesForTab();
        await LoadSettings();
    }

    private async Task LoadIndustries()
    {
        _allIndustries = await _dbService.Industries.Find(_ => true).ToListAsync();

        cmbIndustryFilter.Items.Clear();
        cmbIndustryFilter.Items.Add("All Industries");
        foreach (var industry in _allIndustries)
            cmbIndustryFilter.Items.Add(industry.Name);

        if (cmbIndustryFilter.Items.Count > 0)
            cmbIndustryFilter.SelectedIndex = 0;
    }

    private async Task LoadRecipients()
    {
        _allRecipients = await _dbService.Recipients.Find(_ => true).ToListAsync();
        FilterRecipients();
    }

    private void FilterRecipients()
    {
        listViewRecipients.Items.Clear();

        var selectedIndustry = cmbIndustryFilter.SelectedItem?.ToString();
        var searchText = txtSearch.Text.ToLower();

        var filtered = _allRecipients.AsEnumerable();

        if (selectedIndustry != "All Industries" && !string.IsNullOrEmpty(selectedIndustry))
        {
            var industry = _allIndustries.FirstOrDefault(i => i.Name == selectedIndustry);
            if (industry != null && !string.IsNullOrEmpty(industry.Id))
                filtered = filtered.Where(r => r.Industries.Contains(industry.Id));
        }

        if (!string.IsNullOrWhiteSpace(searchText))
        {
            filtered = filtered.Where(r =>
                r.Email.ToLower().Contains(searchText) ||
                (r.Name?.ToLower().Contains(searchText) ?? false));
        }

        foreach (var recipient in filtered)
        {
            var item = new ListViewItem(recipient.Email);
            item.SubItems.Add(recipient.Name ?? "");
            item.SubItems.Add(string.Join(", ", ResolveIndustryNames(recipient.Industries)));
            item.SubItems.Add(recipient.IsSent ? "Sent" : "Unsent");
            item.SubItems.Add(recipient.CreatedAt.ToLocalTime().ToString("d"));
            item.Tag = recipient;
            listViewRecipients.Items.Add(item);
        }
    }

    private List<string> ResolveIndustryNames(List<string> industryIds)
    {
        var names = new List<string>();
        foreach (var id in industryIds)
        {
            var ind = _allIndustries.FirstOrDefault(i => i.Id == id);
            names.Add(ind?.Name ?? id);
        }
        return names;
    }

    private void LoadAccounts()
    {
        listViewAccounts.Items.Clear();
        _accounts = _accountService.GetAllAccounts();

        if (_accounts.Count == 0)
        {
            lblAccountStatus.Text = "⚠  No Outlook accounts found. Make sure Outlook is open and configured.";
            lblAccountStatus.ForeColor = Color.FromArgb(220, 53, 69);
        }
        else if (_accounts.Count == 1)
        {
            lblAccountStatus.Text = "📧  1 account configured — emails will be sent via this account.";
            lblAccountStatus.ForeColor = Color.FromArgb(13, 110, 253);
        }
        else
        {
            lblAccountStatus.Text = $"🔄  {_accounts.Count} accounts configured — round-robin rotation will be used.";
            lblAccountStatus.ForeColor = Color.FromArgb(25, 135, 84);
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
            if (index % 2 == 0) item.BackColor = Color.FromArgb(248, 249, 250);
            listViewAccounts.Items.Add(item);
            index++;
        }
    }

    private async Task LoadIndustriesForTab()
    {
        listViewIndustries.Items.Clear();
        _industries = await _dbService.Industries.Find(_ => true).ToListAsync();

        foreach (var industry in _industries)
        {
            var item = new ListViewItem(industry.Name);
            item.SubItems.Add(industry.Description ?? "");
            item.SubItems.Add(industry.CreatedAt.ToLocalTime().ToString("d"));
            item.Tag = industry;
            listViewIndustries.Items.Add(item);
        }
    }

    private async Task LoadSettings()
    {
        var settings = await _dbService.Settings.Find(_ => true).FirstOrDefaultAsync();
        if (settings != null)
        {
            txtBatchSize.Text = settings.BatchSize.ToString();
            txtDelay.Text = settings.DelayBetweenBatches.ToString();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  RECIPIENTS EVENTS
    // ═══════════════════════════════════════════════════════════════════════════

    private void BtnAddRecipient_Click(object? sender, EventArgs e)
    {
        var editorForm = new RecipientEditorForm(_dbService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
            _ = LoadRecipients();
    }

    private void BtnEditRecipient_Click(object? sender, EventArgs e)
    {
        if (listViewRecipients.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select a recipient to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var recipient = listViewRecipients.SelectedItems[0].Tag as Recipient;
        if (recipient != null)
        {
            var editorForm = new RecipientEditorForm(_dbService, recipient);
            if (editorForm.ShowDialog() == DialogResult.OK)
                _ = LoadRecipients();
        }
    }

    private async void BtnDeleteRecipient_Click(object? sender, EventArgs e)
    {
        if (listViewRecipients.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select a recipient to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var result = MessageBox.Show("Are you sure you want to delete this recipient?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            var recipient = listViewRecipients.SelectedItems[0].Tag as Recipient;
            if (recipient != null && !string.IsNullOrEmpty(recipient.Id))
            {
                await _dbService.Recipients.DeleteOneAsync(r => r.Id == recipient.Id);
                await LoadRecipients();
            }
        }
    }

    private void BtnImportRecipients_Click(object? sender, EventArgs e)
    {
        var importForm = new ImportRecipientsForm(_dbService);
        if (importForm.ShowDialog() == DialogResult.OK)
            _ = LoadRecipients();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  INDUSTRY EVENTS
    // ═══════════════════════════════════════════════════════════════════════════

    private void BtnAddIndustry_Click(object? sender, EventArgs e)
    {
        var editorForm = new IndustryEditorForm(_dbService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
        {
            _ = LoadIndustriesForTab();
            _ = LoadIndustries(); // also refresh recipient tab filter
        }
    }

    private void BtnEditIndustry_Click(object? sender, EventArgs e)
    {
        if (listViewIndustries.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an industry to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var industry = listViewIndustries.SelectedItems[0].Tag as Industry;
        if (industry != null)
        {
            var editorForm = new IndustryEditorForm(_dbService, industry);
            if (editorForm.ShowDialog() == DialogResult.OK)
            {
                _ = LoadIndustriesForTab();
                _ = LoadIndustries();
            }
        }
    }

    private async void BtnDeleteIndustry_Click(object? sender, EventArgs e)
    {
        if (listViewIndustries.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an industry to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var result = MessageBox.Show("Are you sure you want to delete this industry?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            var industry = listViewIndustries.SelectedItems[0].Tag as Industry;
            if (industry != null && !string.IsNullOrEmpty(industry.Id))
            {
                await _dbService.Industries.DeleteOneAsync(i => i.Id == industry.Id);
                await LoadIndustriesForTab();
                await LoadIndustries();
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  APP SETTINGS EVENTS
    // ═══════════════════════════════════════════════════════════════════════════

    private async void BtnSaveSettings_Click(object? sender, EventArgs e)
    {
        if (!int.TryParse(txtBatchSize.Text, out int batchSize) || batchSize <= 0)
        {
            MessageBox.Show("Please enter a valid batch size.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!int.TryParse(txtDelay.Text, out int delay) || delay < 0)
        {
            MessageBox.Show("Please enter a valid delay time.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var settings = await _dbService.Settings.Find(_ => true).FirstOrDefaultAsync();

        if (settings == null)
        {
            settings = new AppSettings
            {
                BatchSize = batchSize,
                DelayBetweenBatches = delay,
                UpdatedAt = DateTime.UtcNow
            };
            await _dbService.Settings.InsertOneAsync(settings);
        }
        else
        {
            var filter = Builders<AppSettings>.Filter.Eq(s => s.Id, settings.Id);
            var update = Builders<AppSettings>.Update
                .Set(s => s.BatchSize, batchSize)
                .Set(s => s.DelayBetweenBatches, delay)
                .Set(s => s.UpdatedAt, DateTime.UtcNow);
            await _dbService.Settings.UpdateOneAsync(filter, update);
        }

        MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
