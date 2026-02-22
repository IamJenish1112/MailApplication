using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class ImportRecipientsForm : Form
{
    private readonly MongoDbService _dbService;
    private TextBox txtFilePath;
    private Button btnBrowse;
    private CheckedListBox chkIndustries;
    private Button btnImport;
    private Button btnCancel;
    private Label lblStatus;

    private List<Industry> _industries = new();

    public ImportRecipientsForm(MongoDbService dbService)
    {
        _dbService = dbService;
        InitializeComponent();
        SetupUI();
        LoadIndustries();
    }

    private void SetupUI()
    {
        this.Text = "Import Recipients";
        this.Size = new Size(620, 540);
        this.MinimumSize = new Size(500, 450);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = false;

        // === Bottom button panel (must be added to form FIRST for Dock.Bottom to work) ===
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(20, 10, 20, 10)
        };

        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new Size(110, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Dock = DockStyle.Right
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        btnImport = new Button
        {
            Text = "Import",
            Size = new Size(110, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Dock = DockStyle.Right
        };
        btnImport.FlatAppearance.BorderSize = 0;
        btnImport.Click += BtnImport_Click;

        // Add a small spacer between buttons
        var btnSpacer = new Panel { Dock = DockStyle.Right, Width = 10 };

        // Dock.Right items are added in reverse visual order (rightmost first)
        buttonPanel.Controls.Add(btnCancel);
        buttonPanel.Controls.Add(btnSpacer);
        buttonPanel.Controls.Add(btnImport);

        // Add buttonPanel to form FIRST so Dock.Bottom is reserved before Fill takes remaining space
        this.Controls.Add(buttonPanel);

        // === Main content panel (Dock.Fill, added AFTER buttonPanel) ===
        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Padding = new Padding(25, 20, 25, 20)
        };

        var lblFile = new Label
        {
            Text = "Select Text File:",
            Location = new Point(5, 10),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        // File path row: use a container panel with proper anchoring
        var fileRowPanel = new Panel
        {
            Location = new Point(5, 38),
            Size = new Size(540, 38),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        btnBrowse = new Button
        {
            Text = "Browse...",
            Size = new Size(110, 35),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Dock = DockStyle.Right
        };
        btnBrowse.FlatAppearance.BorderSize = 0;
        btnBrowse.Click += BtnBrowse_Click;

        // Small gap between textbox and browse button
        var browseGap = new Panel { Dock = DockStyle.Right, Width = 8 };

        txtFilePath = new TextBox
        {
            ReadOnly = true,
            Font = new Font("Segoe UI", 10),
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(248, 249, 250)
        };

        // Add in correct order for Dock: Right items first, then Fill
        fileRowPanel.Controls.Add(txtFilePath);
        fileRowPanel.Controls.Add(browseGap);
        fileRowPanel.Controls.Add(btnBrowse);

        var lblIndustries = new Label
        {
            Text = "Select Industries:",
            Location = new Point(5, 86),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        chkIndustries = new CheckedListBox
        {
            Location = new Point(5, 114),
            Size = new Size(540, 230),
            Font = new Font("Segoe UI", 10),
            CheckOnClick = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        lblStatus = new Label
        {
            Text = "Select a text file with one email per line",
            Location = new Point(5, 350),
            Size = new Size(540, 40),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        mainPanel.Controls.AddRange(new Control[] { lblFile, fileRowPanel, lblIndustries, chkIndustries, lblStatus });

        // Add mainPanel AFTER buttonPanel so it fills the remaining space
        this.Controls.Add(mainPanel);
    }

    private async void LoadIndustries()
    {
        _industries = await _dbService.Industries.Find(_ => true).ToListAsync();

        chkIndustries.Items.Clear();
        foreach (var industry in _industries)
        {
            chkIndustries.Items.Add(industry.Name);
        }
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var openFileDialog = new OpenFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            Title = "Select Recipients File"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            txtFilePath.Text = openFileDialog.FileName;

            // Count emails for status
            try
            {
                var emailCount = File.ReadAllLines(openFileDialog.FileName)
                    .Count(line => !string.IsNullOrWhiteSpace(line) && line.Contains("@"));
                lblStatus.Text = $"Found {emailCount} email(s) in the file.";
            }
            catch { lblStatus.Text = "File selected. Ready to import."; }
        }
    }

    private async void BtnImport_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(txtFilePath.Text))
        {
            MessageBox.Show("Please select a file to import.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (chkIndustries.CheckedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one industry.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            // Get selected industry ObjectIds
            var selectedIndustryIds = new List<string>();
            for (int i = 0; i < chkIndustries.Items.Count; i++)
            {
                if (chkIndustries.GetItemChecked(i) && i < _industries.Count)
                {
                    var industryId = _industries[i].Id;
                    if (!string.IsNullOrEmpty(industryId))
                    {
                        selectedIndustryIds.Add(industryId);
                    }
                }
            }

            var emails = File.ReadAllLines(txtFilePath.Text)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .Where(email => email.Contains("@"))
                .Distinct()
                .ToList();

            int newCount = 0;
            int updatedCount = 0;

            btnImport.Enabled = false;
            btnImport.Text = "Importing...";
            lblStatus.Text = "Importing...";

            foreach (var email in emails)
            {
                var existing = await _dbService.Recipients.Find(r => r.Email == email).FirstOrDefaultAsync();

                if (existing == null)
                {
                    var newRecipient = new Recipient
                    {
                        Email = email,
                        Industries = selectedIndustryIds,
                        IsSent = false,
                        CreatedAt = DateTime.UtcNow
                    };
                    await _dbService.Recipients.InsertOneAsync(newRecipient);
                    newCount++;
                }
                else
                {
                    // Merge industry ObjectIds (deduplicated)
                    var mergedIndustries = existing.Industries.Union(selectedIndustryIds).Distinct().ToList();
                    var filter = Builders<Recipient>.Filter.Eq(r => r.Id, existing.Id);
                    var update = Builders<Recipient>.Update.Set(r => r.Industries, mergedIndustries);
                    await _dbService.Recipients.UpdateOneAsync(filter, update);
                    updatedCount++;
                }
            }

            MessageBox.Show($"Import completed!\nNew recipients: {newCount}\nUpdated recipients: {updatedCount}",
                "Import Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error importing recipients: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnImport.Enabled = true;
            btnImport.Text = "Import";
        }
    }
}
