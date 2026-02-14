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
        this.Size = new Size(600, 500);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;

        var lblFile = new Label
        {
            Text = "Select Text File:",
            Location = new Point(30, 30),
            Size = new Size(150, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtFilePath = new TextBox
        {
            Location = new Point(30, 60),
            Size = new Size(420, 30),
            ReadOnly = true,
            Font = new Font("Segoe UI", 10)
        };

        btnBrowse = new Button
        {
            Text = "Browse",
            Location = new Point(460, 58),
            Size = new Size(100, 35),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnBrowse.Click += BtnBrowse_Click;

        var lblIndustries = new Label
        {
            Text = "Select Industries:",
            Location = new Point(30, 110),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        chkIndustries = new CheckedListBox
        {
            Location = new Point(30, 140),
            Size = new Size(530, 200),
            Font = new Font("Segoe UI", 10),
            CheckOnClick = true
        };

        lblStatus = new Label
        {
            Text = "Select a text file with one email per line",
            Location = new Point(30, 350),
            Size = new Size(530, 40),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        btnImport = new Button
        {
            Text = "Import",
            Location = new Point(360, 410),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnImport.Click += BtnImport_Click;

        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(470, 410),
            Size = new Size(100, 40),
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        this.Controls.AddRange(new Control[] { lblFile, txtFilePath, btnBrowse, lblIndustries, chkIndustries, lblStatus, btnImport, btnCancel });
    }

    private async void LoadIndustries()
    {
        var industries = await _dbService.Industries.Find(_ => true).ToListAsync();
        
        foreach (var industry in industries)
        {
            chkIndustries.Items.Add(industry.Name);
        }
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var openFileDialog = new OpenFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
            Title = "Select Recipients File"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            txtFilePath.Text = openFileDialog.FileName;
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
            var selectedIndustries = chkIndustries.CheckedItems.Cast<string>().ToList();
            var emails = File.ReadAllLines(txtFilePath.Text)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .Where(email => email.Contains("@"))
                .Distinct()
                .ToList();

            int newCount = 0;
            int updatedCount = 0;

            foreach (var email in emails)
            {
                var existing = await _dbService.Recipients.Find(r => r.Email == email).FirstOrDefaultAsync();

                if (existing == null)
                {
                    var newRecipient = new Recipient
                    {
                        Email = email,
                        Industries = selectedIndustries,
                        IsSent = false,
                        CreatedAt = DateTime.UtcNow
                    };
                    await _dbService.Recipients.InsertOneAsync(newRecipient);
                    newCount++;
                }
                else
                {
                    var updatedIndustries = existing.Industries.Union(selectedIndustries).Distinct().ToList();
                    var filter = Builders<Recipient>.Filter.Eq(r => r.Id, existing.Id);
                    var update = Builders<Recipient>.Update.Set(r => r.Industries, updatedIndustries);
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
    }
}
