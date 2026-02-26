using MailApplication.Models;
using MailApplication.Services;
using MongoDB.Driver;

namespace MailApplication.Forms;

public partial class ImportRecipientsForm : Form
{
    private readonly MongoDbService _dbService;

    // Controls
    private TextBox txtFilePath;
    private Button btnBrowse;
    private CheckedListBox chkIndustries;
    private Button btnImport;
    private Button btnCancel;
    private Label lblStatus;
    private Label lblEmailCount;
    private Panel mainPanel;

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
        // ── Form settings ───────────────────────────────────────────────────
        this.Text = "Import Recipients from File";
        this.Size = new Size(640, 580);
        this.MinimumSize = new Size(560, 520);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // ── Bottom button panel — added FIRST so Dock.Bottom reserves space ─
        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 65,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(0)
        };

        // adjust some left more
        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new Size(120, 40),
            Location = new Point(360, 13),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        btnImport = new Button
        {
            Text = "⬆  Import",
            Size = new Size(120, 40),
            Location = new Point(490, 13),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnImport.FlatAppearance.BorderSize = 0;
        btnImport.Click += BtnImport_Click;

        bottomPanel.Controls.Add(btnCancel);
        bottomPanel.Controls.Add(btnImport);

        this.Controls.Add(bottomPanel);

        // ── Main content panel — added AFTER bottom panel ───────────────────
        mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(0)
        };

        // Title
        var lblTitle = new Label
        {
            Text = "Import Recipients",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            Location = new Point(25, 18),
            Size = new Size(560, 35),
            AutoSize = false
        };

        var lblSubtitle = new Label
        {
            Text = "Select a .txt or .csv file with one email address per line.",
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125),
            Location = new Point(25, 55),
            Size = new Size(560, 22),
            AutoSize = false
        };

        // Divider line
        var divider = new Panel
        {
            Location = new Point(25, 82),
            Size = new Size(570, 1),
            BackColor = Color.FromArgb(222, 226, 230)
        };

        // ── File selection row ───────────────────────────────────────────────
        var lblFile = new Label
        {
            Text = "Text File (.txt / .csv):",
            Location = new Point(25, 100),
            Size = new Size(180, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        // Browse button — explicitly sized & positioned (NOT docked)
        btnBrowse = new Button
        {
            Text = "Browse...",
            Size = new Size(110, 36),
            Location = new Point(480, 128),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnBrowse.FlatAppearance.BorderSize = 0;
        btnBrowse.Click += BtnBrowse_Click;

        // File path textbox — sits to the left of Browse button
        txtFilePath = new TextBox
        {
            Location = new Point(25, 132),
            Size = new Size(444, 30),
            ReadOnly = true,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.FromArgb(248, 249, 250),
            BorderStyle = BorderStyle.FixedSingle,
            ForeColor = Color.FromArgb(73, 80, 87)
        };

        // Email count feedback
        lblEmailCount = new Label
        {
            Text = "",
            Location = new Point(25, 172),
            Size = new Size(560, 22),
            Font = new Font("Segoe UI", 9, FontStyle.Italic),
            ForeColor = Color.FromArgb(25, 135, 84)
        };

        // ── Industry selection ───────────────────────────────────────────────
        var lblIndustries = new Label
        {
            Text = "Assign to Industries:",
            Location = new Point(25, 205),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        var lblIndustriesHint = new Label
        {
            Text = "(Select one or more — same as adding a recipient manually)",
            Location = new Point(230, 207),
            Size = new Size(370, 22),
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        chkIndustries = new CheckedListBox
        {
            Location = new Point(25, 233),
            Size = new Size(568, 200),
            Font = new Font("Segoe UI", 10),
            CheckOnClick = true,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White
        };

        // ── Status label (bottom of content area) ───────────────────────────
        lblStatus = new Label
        {
            Text = "ℹ  No file selected yet. Click Browse to choose a file.",
            Location = new Point(25, 445),
            Size = new Size(568, 40),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        // ── Add all controls to mainPanel ───────────────────────────────────
        mainPanel.Controls.AddRange(new Control[]
        {
            lblTitle, lblSubtitle, divider,
            lblFile,
            txtFilePath, btnBrowse,
            lblEmailCount,
            lblIndustries, lblIndustriesHint,
            chkIndustries,
            lblStatus
        });

        // mainPanel added AFTER bottomPanel
        this.Controls.Add(mainPanel);
    }

    // ── Load industries from MongoDB (same logic as RecipientEditorForm) ────
    private async void LoadIndustries()
    {
        try
        {
            _industries = await _dbService.Industries.Find(_ => true).ToListAsync();

            chkIndustries.Items.Clear();
            foreach (var industry in _industries)
            {
                chkIndustries.Items.Add(industry.Name);
            }

            if (_industries.Count == 0)
            {
                lblStatus.Text = "⚠  No industries found in the database. Please add industries first.";
                lblStatus.ForeColor = Color.FromArgb(220, 53, 69);
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = $"⚠  Failed to load industries: {ex.Message}";
            lblStatus.ForeColor = Color.FromArgb(220, 53, 69);
        }
    }

    // ── Browse for file ──────────────────────────────────────────────────────
    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            Title = "Select Recipients File",
            CheckFileExists = true
        };

        if (dlg.ShowDialog() == DialogResult.OK)
        {
            txtFilePath.Text = dlg.FileName;

            try
            {
                var lines = File.ReadAllLines(dlg.FileName);
                var validEmails = lines
                    .Where(l => !string.IsNullOrWhiteSpace(l) && l.Contains("@"))
                    .Select(l => l.Trim())
                    .Distinct()
                    .Count();

                var totalLines = lines.Length;

                lblEmailCount.Text = $"✔  Found {validEmails} valid email address(es) in {totalLines} line(s).";
                lblEmailCount.ForeColor = Color.FromArgb(25, 135, 84);

                lblStatus.Text = "ℹ  Select one or more industries below, then click Import.";
                lblStatus.ForeColor = Color.FromArgb(108, 117, 125);
            }
            catch (Exception ex)
            {
                lblEmailCount.Text = $"⚠  Could not read file: {ex.Message}";
                lblEmailCount.ForeColor = Color.FromArgb(220, 53, 69);
            }
        }
    }

    // ── Import button — SAME business logic as adding recipients in Recipient
    //    tab: each email stored with selectedIndustryIds, duplicates are merged.
    private async void BtnImport_Click(object? sender, EventArgs e)
    {
        // ── Validations ────────────────────────────────────────────────────
        if (string.IsNullOrWhiteSpace(txtFilePath.Text))
        {
            MessageBox.Show("Please browse and select a file first.", "No File Selected",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (chkIndustries.CheckedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one industry to assign to the imported recipients.",
                "No Industry Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // ── Collect selected industry IDs (ObjectId strings) ───────────────
        var selectedIndustryIds = new List<string>();
        for (int i = 0; i < chkIndustries.Items.Count; i++)
        {
            if (chkIndustries.GetItemChecked(i) && i < _industries.Count)
            {
                var id = _industries[i].Id;
                if (!string.IsNullOrEmpty(id))
                    selectedIndustryIds.Add(id);
            }
        }

        // ── Read and parse emails from file ────────────────────────────────
        List<string> emails;
        try
        {
            emails = File.ReadAllLines(txtFilePath.Text)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .Where(email => email.Contains("@"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading file:\n{ex.Message}", "File Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (emails.Count == 0)
        {
            MessageBox.Show("No valid email addresses found in the selected file.\n\nMake sure each line contains one email address (e.g. user@example.com).",
                "No Emails Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // ── UI: start import ───────────────────────────────────────────────
        btnImport.Enabled = false;
        btnBrowse.Enabled = false;
        btnImport.Text = "Importing...";
        lblStatus.ForeColor = Color.FromArgb(13, 110, 253);
        lblStatus.Text = $"Importing {emails.Count} recipient(s)... please wait.";

        int newCount = 0;
        int updatedCount = 0;
        int errorCount = 0;

        try
        {
            foreach (var email in emails)
            {
                try
                {
                    // Same logic as RecipientEditorForm: check for duplicate, merge industries
                    var existing = await _dbService.Recipients
                        .Find(r => r.Email == email)
                        .FirstOrDefaultAsync();

                    if (existing == null)
                    {
                        // NEW recipient — same fields as adding manually
                        var newRecipient = new Recipient
                        {
                            Email = email,
                            Name = null,                      // no name from .txt file
                            Industries = selectedIndustryIds, // assigned industry IDs
                            IsSent = false,
                            CreatedAt = DateTime.UtcNow
                        };
                        await _dbService.Recipients.InsertOneAsync(newRecipient);
                        newCount++;
                    }
                    else
                    {
                        // EXISTING recipient — merge industry IDs (deduplicated), same as edit
                        var mergedIndustries = existing.Industries
                            .Union(selectedIndustryIds)
                            .Distinct()
                            .ToList();

                        var filter = Builders<Recipient>.Filter.Eq(r => r.Id, existing.Id);
                        var update = Builders<Recipient>.Update
                            .Set(r => r.Industries, mergedIndustries);
                        await _dbService.Recipients.UpdateOneAsync(filter, update);
                        updatedCount++;
                    }
                }
                catch
                {
                    errorCount++;
                }
            }

            var resultMsg = $"Import completed successfully!\n\n" +
                            $"✔  New recipients added : {newCount}\n" +
                            $"✔  Existing updated      : {updatedCount}";
            if (errorCount > 0)
                resultMsg += $"\n⚠  Errors skipped       : {errorCount}";

            MessageBox.Show(resultMsg, "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.FromArgb(220, 53, 69);
            lblStatus.Text = $"⚠  Import failed: {ex.Message}";
            MessageBox.Show($"Import failed:\n{ex.Message}", "Import Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnImport.Enabled = true;
            btnBrowse.Enabled = true;
            btnImport.Text = "⬆  Import";
        }
    }
}
