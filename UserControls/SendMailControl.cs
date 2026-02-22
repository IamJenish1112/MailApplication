using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class SendMailControl : UserControl
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;
    private readonly EmailSendingService _emailService;

    private ComboBox cmbIndustry;
    private ComboBox cmbTemplateSource;
    private ComboBox cmbDraft;
    private ComboBox cmbSenderAccount;
    private ListView listViewRecipients;
    private CheckBox chkSelectAll;
    private TextBox txtBatchSize;
    private TextBox txtDelay;
    private Button btnPreviewDraft;
    private Button btnImportRecipients;
    private Button btnMarkSent;
    private Button btnMarkUnsent;
    private Button btnStartSending;
    private Button btnStopSending;
    private ProgressBar progressBar;
    private Label lblProgress;
    private RichTextBox txtLog;

    private List<Recipient> _allRecipients = new();
    private List<Draft> _appDrafts = new();
    private List<Draft> _outlookDrafts = new();
    private List<Draft> _currentDrafts = new();
    private List<Industry> _allIndustries = new();
    private List<EmailAccount> _outlookAccounts = new();
    private CancellationTokenSource? _cancellationTokenSource;

    public SendMailControl(MongoDbService dbService, OutlookService outlookService, EmailSendingService emailService)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        _emailService = emailService;
        InitializeComponent();
        SetupUI();
        LoadData();
        SetupEventHandlers();
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);
        this.AutoScroll = true;

        var titleLabel = new Label
        {
            Text = "Send Bulk Mail",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        // Row 1: Industry + Template Source
        var lblIndustry = new Label
        {
            Text = "Select Industry:",
            Location = new Point(0, 50),
            Size = new Size(130, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbIndustry = new ComboBox
        {
            Location = new Point(140, 50),
            Size = new Size(220, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        var lblTemplateSource = new Label
        {
            Text = "Template Source:",
            Location = new Point(380, 50),
            Size = new Size(130, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbTemplateSource = new ComboBox
        {
            Location = new Point(520, 50),
            Size = new Size(180, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };
        cmbTemplateSource.Items.AddRange(new object[] { "Application Templates", "Outlook Drafts" });
        cmbTemplateSource.SelectedIndex = 0;

        // Row 2: Draft selection + Sender Account
        var lblDraft = new Label
        {
            Text = "Select Template:",
            Location = new Point(0, 95),
            Size = new Size(130, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbDraft = new ComboBox
        {
            Location = new Point(140, 95),
            Size = new Size(300, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        btnPreviewDraft = new Button
        {
            Text = "Preview",
            Location = new Point(450, 93),
            Size = new Size(90, 33),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnPreviewDraft.FlatAppearance.BorderSize = 0;

        var lblSender = new Label
        {
            Text = "Send As:",
            Location = new Point(560, 95),
            Size = new Size(80, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbSenderAccount = new ComboBox
        {
            Location = new Point(650, 95),
            Size = new Size(320, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        // Select All checkbox
        chkSelectAll = new CheckBox
        {
            Text = "Select All Recipients",
            Location = new Point(0, 140),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        // Recipients ListView
        listViewRecipients = new ListView
        {
            Location = new Point(0, 170),
            Size = new Size(980, 270),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            CheckBoxes = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewRecipients.Columns.Add("Email", 280);
        listViewRecipients.Columns.Add("Name", 170);
        listViewRecipients.Columns.Add("Industries", 230);
        listViewRecipients.Columns.Add("Status", 100);
        listViewRecipients.Columns.Add("Last Sent", 130);

        // Button panel
        var buttonPanel = new Panel
        {
            Location = new Point(0, 450),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnImportRecipients = CreateButton("Import Recipients", 0);
        btnMarkSent = CreateButton("Mark as Sent", 160);
        btnMarkUnsent = CreateButton("Mark as Unsent", 320);

        buttonPanel.Controls.AddRange(new Control[] { btnImportRecipients, btnMarkSent, btnMarkUnsent });

        // Batch Size & Delay
        var lblBatchSize = new Label
        {
            Text = "Batch Size (BCC):",
            Location = new Point(0, 515),
            Size = new Size(140, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtBatchSize = new TextBox
        {
            Location = new Point(150, 515),
            Size = new Size(80, 30),
            Text = "50",
            Font = new Font("Segoe UI", 10)
        };

        var lblDelay = new Label
        {
            Text = "Delay (seconds):",
            Location = new Point(260, 515),
            Size = new Size(130, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtDelay = new TextBox
        {
            Location = new Point(400, 515),
            Size = new Size(80, 30),
            Text = "60",
            Font = new Font("Segoe UI", 10)
        };

        btnStartSending = new Button
        {
            Text = "▶ Start Sending",
            Location = new Point(520, 510),
            Size = new Size(150, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnStartSending.FlatAppearance.BorderSize = 0;

        btnStopSending = new Button
        {
            Text = "⏹ Stop",
            Location = new Point(680, 510),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Enabled = false
        };
        btnStopSending.FlatAppearance.BorderSize = 0;

        progressBar = new ProgressBar
        {
            Location = new Point(0, 565),
            Size = new Size(980, 25),
            Style = ProgressBarStyle.Continuous
        };

        lblProgress = new Label
        {
            Text = "Ready to send",
            Location = new Point(0, 597),
            Size = new Size(980, 25),
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        txtLog = new RichTextBox
        {
            Location = new Point(0, 625),
            Size = new Size(980, 140),
            ReadOnly = true,
            Font = new Font("Consolas", 9),
            BackColor = Color.FromArgb(248, 249, 250)
        };

        this.Controls.AddRange(new Control[] {
            titleLabel,
            lblIndustry, cmbIndustry,
            lblTemplateSource, cmbTemplateSource,
            lblDraft, cmbDraft, btnPreviewDraft,
            lblSender, cmbSenderAccount,
            chkSelectAll,
            listViewRecipients, buttonPanel,
            lblBatchSize, txtBatchSize, lblDelay, txtDelay,
            btnStartSending, btnStopSending,
            progressBar, lblProgress, txtLog
        });
    }

    private Button CreateButton(string text, int left)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(left, 5),
            Size = new Size(150, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    private void SetupEventHandlers()
    {
        cmbIndustry.SelectedIndexChanged += (s, e) => FilterRecipients();
        cmbTemplateSource.SelectedIndexChanged += (s, e) => RefreshDraftList();
        chkSelectAll.CheckedChanged += ChkSelectAll_CheckedChanged;
        btnPreviewDraft.Click += BtnPreviewDraft_Click;
        btnImportRecipients.Click += BtnImportRecipients_Click;
        btnMarkSent.Click += BtnMarkSent_Click;
        btnMarkUnsent.Click += BtnMarkUnsent_Click;
        btnStartSending.Click += BtnStartSending_Click;
        btnStopSending.Click += BtnStopSending_Click;

        _emailService.BatchProgress += EmailService_BatchProgress;
        _emailService.StatusUpdate += EmailService_StatusUpdate;
    }

    private async void LoadData()
    {
        await LoadIndustries();
        await LoadSettings(); // Auto-load batch size & delay from Settings
        await LoadDrafts();
        await LoadRecipients();
        LoadSenderAccounts();
    }

    /// <summary>
    /// Auto-loads batch size and delay from Settings collection in MongoDB.
    /// </summary>
    private async Task LoadSettings()
    {
        try
        {
            var settings = await _dbService.Settings.Find(_ => true).FirstOrDefaultAsync();
            if (settings != null)
            {
                txtBatchSize.Text = settings.BatchSize.ToString();
                txtDelay.Text = settings.DelayBetweenBatches.ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading settings: {ex.Message}");
        }
    }

    private async Task LoadIndustries()
    {
        _allIndustries = await _dbService.Industries.Find(_ => true).ToListAsync();
        cmbIndustry.Items.Clear();
        cmbIndustry.Items.Add("All Industries");

        foreach (var industry in _allIndustries)
        {
            cmbIndustry.Items.Add(industry.Name);
        }

        if (cmbIndustry.Items.Count > 0)
            cmbIndustry.SelectedIndex = 0;
    }

    private async Task LoadDrafts()
    {
        _appDrafts = await _dbService.Drafts.Find(_ => true).ToListAsync();

        if (_outlookService.IsAvailable)
        {
            try { _outlookDrafts = _outlookService.GetDraftsFromOutlook(); }
            catch { _outlookDrafts = new List<Draft>(); }
        }

        RefreshDraftList();
    }

    private void RefreshDraftList()
    {
        cmbDraft.Items.Clear();

        if (cmbTemplateSource.SelectedIndex == 0) // Application Templates
        {
            _currentDrafts = _appDrafts;
        }
        else // Outlook Drafts
        {
            _currentDrafts = _outlookDrafts;
        }

        foreach (var draft in _currentDrafts)
        {
            cmbDraft.Items.Add($"{draft.Subject} [{draft.Source}]");
        }

        if (cmbDraft.Items.Count > 0)
            cmbDraft.SelectedIndex = 0;
    }

    private void LoadSenderAccounts()
    {
        cmbSenderAccount.Items.Clear();

        if (_outlookService.IsAvailable)
        {
            _outlookAccounts = _outlookService.GetOutlookAccounts();

            foreach (var account in _outlookAccounts)
            {
                cmbSenderAccount.Items.Add($"{account.AccountName} ({account.SmtpAddress})");
            }
        }

        if (cmbSenderAccount.Items.Count > 0)
        {
            // Select the default account
            var defaultIndex = _outlookAccounts.FindIndex(a => a.IsDefault);
            cmbSenderAccount.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
        }
        else
        {
            cmbSenderAccount.Items.Add("No accounts available");
            cmbSenderAccount.SelectedIndex = 0;
        }
    }

    private async Task LoadRecipients()
    {
        _allRecipients = await _dbService.Recipients.Find(_ => true).ToListAsync();
        FilterRecipients();
    }

    private void FilterRecipients()
    {
        listViewRecipients.Items.Clear();

        var selectedIndustry = cmbIndustry.SelectedItem?.ToString();

        IEnumerable<Recipient> filteredRecipients;

        if (selectedIndustry == "All Industries" || string.IsNullOrEmpty(selectedIndustry))
        {
            filteredRecipients = _allRecipients;
        }
        else
        {
            // Find industry by name to get its ObjectId
            var industry = _allIndustries.FirstOrDefault(i => i.Name == selectedIndustry);
            if (industry != null && !string.IsNullOrEmpty(industry.Id))
            {
                filteredRecipients = _allRecipients.Where(r => r.Industries.Contains(industry.Id));
            }
            else
            {
                filteredRecipients = _allRecipients;
            }
        }

        foreach (var recipient in filteredRecipients)
        {
            var item = new ListViewItem(recipient.Email);
            item.SubItems.Add(recipient.Name ?? "");

            // Resolve industry IDs to names
            var industryNames = ResolveIndustryNames(recipient.Industries);
            item.SubItems.Add(string.Join(", ", industryNames));
            item.SubItems.Add(recipient.IsSent ? "Sent" : "Unsent");
            item.SubItems.Add(recipient.LastSentAt?.ToLocalTime().ToString("g") ?? "Never");
            item.Tag = recipient;

            if (!recipient.IsSent)
                item.Checked = true;

            listViewRecipients.Items.Add(item);
        }

        // Update Select All checkbox state
        UpdateSelectAllState();
    }

    private List<string> ResolveIndustryNames(List<string> industryIds)
    {
        var names = new List<string>();
        foreach (var id in industryIds)
        {
            var industry = _allIndustries.FirstOrDefault(i => i.Id == id);
            names.Add(industry?.Name ?? id);
        }
        return names;
    }

    private void ChkSelectAll_CheckedChanged(object? sender, EventArgs e)
    {
        foreach (ListViewItem item in listViewRecipients.Items)
        {
            item.Checked = chkSelectAll.Checked;
        }
    }

    private void UpdateSelectAllState()
    {
        if (listViewRecipients.Items.Count == 0) return;

        bool allChecked = true;
        foreach (ListViewItem item in listViewRecipients.Items)
        {
            if (!item.Checked) { allChecked = false; break; }
        }
        chkSelectAll.Checked = allChecked;
    }

    private void BtnPreviewDraft_Click(object? sender, EventArgs e)
    {
        if (cmbDraft.SelectedIndex < 0 || cmbDraft.SelectedIndex >= _currentDrafts.Count) return;

        var draft = _currentDrafts[cmbDraft.SelectedIndex];
        var previewForm = new DraftPreviewForm(draft);
        previewForm.ShowDialog();
    }

    private async void BtnImportRecipients_Click(object? sender, EventArgs e)
    {
        var importForm = new ImportRecipientsForm(_dbService);
        if (importForm.ShowDialog() == DialogResult.OK)
        {
            await LoadRecipients();
        }
    }

    private async void BtnMarkSent_Click(object? sender, EventArgs e)
    {
        await UpdateSelectedRecipientsStatus(true);
    }

    private async void BtnMarkUnsent_Click(object? sender, EventArgs e)
    {
        await UpdateSelectedRecipientsStatus(false);
    }

    private async Task UpdateSelectedRecipientsStatus(bool isSent)
    {
        var checkedCount = listViewRecipients.CheckedItems.Count;
        if (checkedCount == 0)
        {
            MessageBox.Show("Please select at least one recipient.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        foreach (ListViewItem item in listViewRecipients.CheckedItems)
        {
            var recipient = item.Tag as Recipient;
            if (recipient != null && !string.IsNullOrEmpty(recipient.Id))
            {
                var filter = Builders<Recipient>.Filter.Eq(r => r.Id, recipient.Id);
                var update = Builders<Recipient>.Update.Set(r => r.IsSent, isSent);
                await _dbService.Recipients.UpdateOneAsync(filter, update);
            }
        }

        await LoadRecipients();
    }

    private async void BtnStartSending_Click(object? sender, EventArgs e)
    {
        if (cmbDraft.SelectedIndex < 0 || cmbDraft.SelectedIndex >= _currentDrafts.Count)
        {
            MessageBox.Show("Please select a draft template.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var checkedRecipients = listViewRecipients.CheckedItems.Cast<ListViewItem>()
            .Select(item => item.Tag as Recipient)
            .Where(r => r != null)
            .Cast<Recipient>()
            .ToList();

        if (checkedRecipients.Count == 0)
        {
            MessageBox.Show("Please select at least one recipient.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!int.TryParse(txtBatchSize.Text, out int batchSize) || batchSize <= 0)
        {
            MessageBox.Show("Please enter a valid batch size.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!int.TryParse(txtDelay.Text, out int delay) || delay < 0)
        {
            MessageBox.Show("Please enter a valid delay time.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var draft = _currentDrafts[cmbDraft.SelectedIndex];

        // Get selected sender account SMTP address
        string? senderSmtp = null;
        if (cmbSenderAccount.SelectedIndex >= 0 && cmbSenderAccount.SelectedIndex < _outlookAccounts.Count)
        {
            senderSmtp = _outlookAccounts[cmbSenderAccount.SelectedIndex].SmtpAddress;
        }

        btnStartSending.Enabled = false;
        btnStopSending.Enabled = true;
        progressBar.Value = 0;
        txtLog.Clear();

        _cancellationTokenSource = new CancellationTokenSource();

        try
        {
            await _emailService.SendBulkEmailsAsync(
                checkedRecipients, draft, batchSize, delay,
                senderSmtp, _cancellationTokenSource.Token);
        }
        catch (Exception ex)
        {
            LogMessage($"Error: {ex.Message}");
        }
        finally
        {
            btnStartSending.Enabled = true;
            btnStopSending.Enabled = false;
            await LoadRecipients();
        }
    }

    private void BtnStopSending_Click(object? sender, EventArgs e)
    {
        _cancellationTokenSource?.Cancel();
        btnStopSending.Enabled = false;
    }

    private void EmailService_BatchProgress(object? sender, BatchProgressEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(() => EmailService_BatchProgress(sender, e));
            return;
        }

        progressBar.Maximum = e.TotalBatches;
        progressBar.Value = e.CurrentBatch;
        lblProgress.Text = $"Batch {e.CurrentBatch}/{e.TotalBatches} - {e.EmailsSent} emails sent";
    }

    private void EmailService_StatusUpdate(object? sender, string message)
    {
        if (InvokeRequired)
        {
            Invoke(() => EmailService_StatusUpdate(sender, message));
            return;
        }

        LogMessage(message);
    }

    private void LogMessage(string message)
    {
        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
        txtLog.ScrollToCaret();
    }
}
