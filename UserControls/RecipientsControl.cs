using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class RecipientsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private ComboBox cmbIndustry;
    private ListView listViewRecipients;
    private Button btnAddRecipient;
    private Button btnEditRecipient;
    private Button btnDeleteRecipient;
    private Button btnImport;
    private Button btnRefresh;
    private TextBox txtSearch;

    private List<Recipient> _allRecipients = new();
    private List<Industry> _allIndustries = new();

    public RecipientsControl(MongoDbService dbService)
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

        var titleLabel = new Label
        {
            Text = "Recipient Management",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        var lblIndustry = new Label
        {
            Text = "Filter by Industry:",
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
        cmbIndustry.SelectedIndexChanged += (s, e) => FilterRecipients();

        var lblSearch = new Label
        {
            Text = "Search:",
            Location = new Point(450, 60),
            Size = new Size(80, 25),
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        txtSearch = new TextBox
        {
            Location = new Point(540, 60),
            Size = new Size(250, 30),
            Font = new Font("Segoe UI", 10)
        };
        txtSearch.TextChanged += (s, e) => FilterRecipients();

        listViewRecipients = new ListView
        {
            Location = new Point(0, 110),
            Size = new Size(980, 450),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewRecipients.Columns.Add("Email", 300);
        listViewRecipients.Columns.Add("Name", 200);
        listViewRecipients.Columns.Add("Industries", 280);
        listViewRecipients.Columns.Add("Status", 100);
        listViewRecipients.Columns.Add("Created", 100);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 580),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnAddRecipient = CreateButton("Add Recipient", 0);
        btnEditRecipient = CreateButton("Edit", 160);
        btnDeleteRecipient = CreateButton("Delete", 280);
        btnImport = CreateButton("Import from File", 400);
        btnRefresh = CreateButton("Refresh", 570);

        btnAddRecipient.Click += BtnAddRecipient_Click;
        btnEditRecipient.Click += BtnEditRecipient_Click;
        btnDeleteRecipient.Click += BtnDeleteRecipient_Click;
        btnImport.Click += BtnImport_Click;
        btnRefresh.Click += (s, e) => LoadData();

        buttonPanel.Controls.AddRange(new Control[] { btnAddRecipient, btnEditRecipient, btnDeleteRecipient, btnImport, btnRefresh });

        this.Controls.AddRange(new Control[] { titleLabel, lblIndustry, cmbIndustry, lblSearch, txtSearch, listViewRecipients, buttonPanel });
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

    private async void LoadData()
    {
        await LoadIndustries();
        await LoadRecipients();
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

    private async Task LoadRecipients()
    {
        _allRecipients = await _dbService.Recipients.Find(_ => true).ToListAsync();
        FilterRecipients();
    }

    private void FilterRecipients()
    {
        listViewRecipients.Items.Clear();

        var selectedIndustry = cmbIndustry.SelectedItem?.ToString();
        var searchText = txtSearch.Text.ToLower();

        var filteredRecipients = _allRecipients.AsEnumerable();

        if (selectedIndustry != "All Industries" && !string.IsNullOrEmpty(selectedIndustry))
        {
            // Find industry by name to get its ObjectId
            var industry = _allIndustries.FirstOrDefault(i => i.Name == selectedIndustry);
            if (industry != null && !string.IsNullOrEmpty(industry.Id))
            {
                filteredRecipients = filteredRecipients.Where(r => r.Industries.Contains(industry.Id));
            }
        }

        if (!string.IsNullOrWhiteSpace(searchText))
        {
            filteredRecipients = filteredRecipients.Where(r =>
                r.Email.ToLower().Contains(searchText) ||
                (r.Name?.ToLower().Contains(searchText) ?? false));
        }

        foreach (var recipient in filteredRecipients)
        {
            var item = new ListViewItem(recipient.Email);
            item.SubItems.Add(recipient.Name ?? "");

            // Resolve industry ObjectIds to names
            var industryNames = ResolveIndustryNames(recipient.Industries);
            item.SubItems.Add(string.Join(", ", industryNames));
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
            var industry = _allIndustries.FirstOrDefault(i => i.Id == id);
            names.Add(industry?.Name ?? id);
        }
        return names;
    }

    private void BtnAddRecipient_Click(object? sender, EventArgs e)
    {
        var editorForm = new RecipientEditorForm(_dbService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
        {
            LoadData();
        }
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
            {
                LoadData();
            }
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
                LoadData();
            }
        }
    }

    private void BtnImport_Click(object? sender, EventArgs e)
    {
        var importForm = new ImportRecipientsForm(_dbService);
        if (importForm.ShowDialog() == DialogResult.OK)
        {
            LoadData();
        }
    }
}
