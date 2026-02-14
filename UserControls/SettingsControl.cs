using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class SettingsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private ListView listViewIndustries;
    private Button btnAddIndustry;
    private Button btnEditIndustry;
    private Button btnDeleteIndustry;
    private Button btnRefresh;
    private TextBox txtBatchSize;
    private TextBox txtDelay;
    private Button btnSaveSettings;

    private List<Industry> _industries = new();

    public SettingsControl(MongoDbService dbService)
    {
        _dbService = dbService;
        InitializeComponent();
        SetupUI();
        LoadData();
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);
        this.AutoScroll = true;

        var titleLabel = new Label
        {
            Text = "Settings",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        var lblIndustries = new Label
        {
            Text = "Industry Management",
            Location = new Point(0, 60),
            Size = new Size(300, 30),
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        listViewIndustries = new ListView
        {
            Location = new Point(0, 100),
            Size = new Size(980, 300),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewIndustries.Columns.Add("Industry Name", 500);
        listViewIndustries.Columns.Add("Description", 350);
        listViewIndustries.Columns.Add("Created", 130);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 420),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnAddIndustry = CreateButton("Add Industry", 0);
        btnEditIndustry = CreateButton("Edit", 150);
        btnDeleteIndustry = CreateButton("Delete", 270);
        btnRefresh = CreateButton("Refresh", 390);

        btnAddIndustry.Click += BtnAddIndustry_Click;
        btnEditIndustry.Click += BtnEditIndustry_Click;
        btnDeleteIndustry.Click += BtnDeleteIndustry_Click;
        btnRefresh.Click += (s, e) => LoadData();

        buttonPanel.Controls.AddRange(new Control[] { btnAddIndustry, btnEditIndustry, btnDeleteIndustry, btnRefresh });

        var lblAppSettings = new Label
        {
            Text = "Application Settings",
            Location = new Point(0, 500),
            Size = new Size(300, 30),
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        var lblBatchSize = new Label
        {
            Text = "Default Batch Size:",
            Location = new Point(0, 550),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtBatchSize = new TextBox
        {
            Location = new Point(210, 550),
            Size = new Size(150, 30),
            Text = "50",
            Font = new Font("Segoe UI", 10)
        };

        var lblDelay = new Label
        {
            Text = "Default Delay (seconds):",
            Location = new Point(400, 550),
            Size = new Size(200, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtDelay = new TextBox
        {
            Location = new Point(610, 550),
            Size = new Size(150, 30),
            Text = "60",
            Font = new Font("Segoe UI", 10)
        };

        btnSaveSettings = new Button
        {
            Text = "Save Settings",
            Location = new Point(0, 610),
            Size = new Size(150, 40),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSaveSettings.Click += BtnSaveSettings_Click;

        this.Controls.AddRange(new Control[] {
            titleLabel, lblIndustries, listViewIndustries, buttonPanel,
            lblAppSettings, lblBatchSize, txtBatchSize, lblDelay, txtDelay, btnSaveSettings
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

    private async void LoadData()
    {
        await LoadIndustries();
        await LoadSettings();
    }

    private async Task LoadIndustries()
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

    private void BtnAddIndustry_Click(object? sender, EventArgs e)
    {
        var editorForm = new IndustryEditorForm(_dbService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
        {
            LoadData();
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
                LoadData();
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
                LoadData();
            }
        }
    }

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
