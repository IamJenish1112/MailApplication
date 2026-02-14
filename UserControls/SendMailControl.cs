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
    private ComboBox cmbDraft;
    private ListView listViewRecipients;
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
    private List<Draft> _allDrafts = new();
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

        var lblIndustry = new Label
        {
            Text = "Select Industry:",
            Location = new Point(0, 60),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbIndustry = new ComboBox
        {
            Location = new Point(160, 60),
            Size = new Size(250, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        var lblDraft = new Label
        {
            Text = "Select Draft:",
            Location = new Point(450, 60),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        cmbDraft = new ComboBox
        {
            Location = new Point(610, 60),
            Size = new Size(250, 30),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        btnPreviewDraft = new Button
        {
            Text = "Preview",
            Location = new Point(870, 58),
            Size = new Size(100, 35),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        listViewRecipients = new ListView
        {
            Location = new Point(0, 110),
            Size = new Size(980, 300),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            CheckBoxes = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewRecipients.Columns.Add("Email", 300);
        listViewRecipients.Columns.Add("Name", 200);
        listViewRecipients.Columns.Add("Industries", 250);
        listViewRecipients.Columns.Add("Status", 100);
        listViewRecipients.Columns.Add("Last Sent", 130);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 420),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnImportRecipients = CreateButton("Import Recipients", 0);
        btnMarkSent = CreateButton("Mark as Sent", 150);
        btnMarkUnsent = CreateButton("Mark as Unsent", 300);

        buttonPanel.Controls.AddRange(new Control[] { btnImportRecipients, btnMarkSent, btnMarkUnsent });

        var lblBatchSize = new Label
        {
            Text = "Batch Size (BCC):",
            Location = new Point(0, 490),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtBatchSize = new TextBox
        {
            Location = new Point(160, 490),
            Size = new Size(100, 30),
            Text = "50",
            Font = new Font("Segoe UI", 10)
        };

        var lblDelay = new Label
        {
            Text = "Delay (seconds):",
            Location = new Point(300, 490),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtDelay = new TextBox
        {
            Location = new Point(460, 490),
            Size = new Size(100, 30),
            Text = "60",
            Font = new Font("Segoe UI", 10)
        };

        btnStartSending = new Button
        {
            Text = "Start Sending",
            Location = new Point(600, 485),
            Size = new Size(140, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Cursor = Cursors.Hand
        };

        btnStopSending = new Button
        {
            Text = "Stop",
            Location = new Point(750, 485),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Enabled = false
        };

        progressBar = new ProgressBar
        {
            Location = new Point(0, 550),
            Size = new Size(980, 30),
            Style = ProgressBarStyle.Continuous
        };

        lblProgress = new Label
        {
            Text = "Ready to send",
            Location = new Point(0, 590),
            Size = new Size(980, 25),
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        txtLog = new RichTextBox
        {
            Location = new Point(0, 625),
            Size = new Size(980, 150),
            ReadOnly = true,
            Font = new Font("Consolas", 9),
            BackColor = Color.FromArgb(248, 249, 250)
        };

        this.Controls.AddRange(new Control[] {
            titleLabel, lblIndustry, cmbIndustry, lblDraft, cmbDraft, btnPreviewDraft,
            listViewRecipients, buttonPanel, lblBatchSize, txtBatchSize, lblDelay, txtDelay,
            btnStartSending, btnStopSending, progressBar, lblProgress, txtLog
        });
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

    private void SetupEventHandlers()
    {
        cmbIndustry.SelectedIndexChanged += (s, e) => FilterRecipients();
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
        await LoadDrafts();
        await LoadRecipients();
    }

    private async Task LoadIndustries()
    {
        var industries = await _dbService.Industries.Find(_ => true).ToListAsync();
        cmbIndustry.Items.Clear();
        cmbIndustry.Items.Add("All Industries");

        foreach (var industry in industries)
        {
            cmbIndustry.Items.Add(industry.Name);
        }

        if (cmbIndustry.Items.Count > 0)
            cmbIndustry.SelectedIndex = 0;
    }

    private async Task LoadDrafts()
    {
        _allDrafts = await _dbService.Drafts.Find(_ => true).ToListAsync();
        cmbDraft.Items.Clear();

        foreach (var draft in _allDrafts)
        {
            cmbDraft.Items.Add(draft.Subject);
        }

        if (cmbDraft.Items.Count > 0)
            cmbDraft.SelectedIndex = 0;
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
        var filteredRecipients = selectedIndustry == "All Industries" || string.IsNullOrEmpty(selectedIndustry)
            ? _allRecipients
            : _allRecipients.Where(r => r.Industries.Contains(selectedIndustry)).ToList();

        foreach (var recipient in filteredRecipients)
        {
            var item = new ListViewItem(recipient.Email);
            item.SubItems.Add(recipient.Name ?? "");
            item.SubItems.Add(string.Join(", ", recipient.Industries));
            item.SubItems.Add(recipient.IsSent ? "Sent" : "Unsent");
            item.SubItems.Add(recipient.LastSentAt?.ToLocalTime().ToString("g") ?? "Never");
            item.Tag = recipient;

            if (!recipient.IsSent)
                item.Checked = true;

            listViewRecipients.Items.Add(item);
        }
    }

    private void BtnPreviewDraft_Click(object? sender, EventArgs e)
    {
        if (cmbDraft.SelectedIndex < 0) return;

        var draft = _allDrafts[cmbDraft.SelectedIndex];
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
        if (cmbDraft.SelectedIndex < 0)
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

        var draft = _allDrafts[cmbDraft.SelectedIndex];

        btnStartSending.Enabled = false;
        btnStopSending.Enabled = true;
        progressBar.Value = 0;
        txtLog.Clear();

        _cancellationTokenSource = new CancellationTokenSource();

        try
        {
            await _emailService.SendBulkEmailsAsync(checkedRecipients, draft, batchSize, delay, _cancellationTokenSource.Token);
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
