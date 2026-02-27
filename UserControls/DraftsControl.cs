using MailApplication.Models;
using MailApplication.Services;
using MailApplication.Forms;
using MongoDB.Driver;

namespace MailApplication.UserControls;

public partial class DraftsControl : UserControl
{
    private readonly MongoDbService _dbService;
    private readonly OutlookService _outlookService;

    private TabControl tabControl;
    private TabPage tabAppDrafts;
    private TabPage tabOutlookDrafts;

    private ListView listViewApp;
    private ListView listViewOutlook;
    private Button btnCreateNew;
    private Button btnEdit;
    private Button btnDelete;
    private Button btnRefresh;
    private Button btnPreview;

    private List<Draft> _appDrafts = new();
    private List<Draft> _outlookDrafts = new();

    public DraftsControl(MongoDbService dbService, OutlookService outlookService)
    {
        _dbService = dbService;
        _outlookService = outlookService;
        InitializeComponent();
        SetupUI();
        LoadDrafts();
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);

        var titleLabel = new Label
        {
            Text = "Drafts Management",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        tabControl = new TabControl
        {
            Location = new Point(0, 60),
            Size = new Size(this.Width - 40, this.Height - 180),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            Font = new Font("Segoe UI", 10)
        };

        tabAppDrafts = new TabPage("Application Drafts (In DB)");
        tabOutlookDrafts = new TabPage("Outlook Drafts");

        listViewApp = CreateListView();
        listViewOutlook = CreateListView();

        tabAppDrafts.Controls.Add(listViewApp);
        tabOutlookDrafts.Controls.Add(listViewOutlook);

        tabControl.TabPages.Add(tabAppDrafts);
        tabControl.TabPages.Add(tabOutlookDrafts);

        var buttonPanel = new Panel
        {
            Height = 60,
            Dock = DockStyle.Bottom,
            BackColor = Color.White,
            Padding = new Padding(0, 10, 0, 0)
        };

        btnCreateNew = CreateButton("Create New (Compose)", 0);
        btnEdit = CreateButton("Edit", 140);
        btnDelete = CreateButton("Delete", 280);
        btnPreview = CreateButton("Preview", 420);
        btnRefresh = CreateButton("Refresh", 560);

        btnCreateNew.Click += BtnCreateNew_Click;
        btnEdit.Click += BtnEdit_Click;
        btnDelete.Click += BtnDelete_Click;
        btnPreview.Click += BtnPreview_Click;
        btnRefresh.Click += (s, e) => LoadDrafts();

        // Disable Outlook tab if Outlook is not available
        if (!_outlookService.IsAvailable)
        {
            tabOutlookDrafts.Enabled = false;
        }

        buttonPanel.Controls.AddRange(new Control[] { btnCreateNew, btnEdit, btnDelete, btnPreview, btnRefresh });

        this.Controls.Add(titleLabel);
        this.Controls.Add(buttonPanel);
        this.Controls.Add(tabControl);
    }

    private ListView CreateListView()
    {
        var listView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.White
        };

        listView.Columns.Add("Subject", 400);
        listView.Columns.Add("Created", 200);
        listView.Columns.Add("Updated", 200);
        listView.Columns.Add("Source", 150);

        return listView;
    }

    private Button CreateButton(string text, int left)
    {
        return new Button
        {
            Text = text,
            Location = new Point(left, 0),
            Size = new Size(130, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
    }

    private async void LoadDrafts()
    {
        await LoadAppDrafts();
        LoadOutlookDrafts();
    }

    private async Task LoadAppDrafts()
    {
        listViewApp.Items.Clear();
        _appDrafts = await _dbService.Drafts.Find(d => d.Source == "Application").ToListAsync();

        foreach (var draft in _appDrafts)
        {
            var item = new ListViewItem(draft.Subject);
            item.SubItems.Add(draft.CreatedAt.ToLocalTime().ToString("g"));
            item.SubItems.Add(draft.UpdatedAt.ToLocalTime().ToString("g"));
            item.SubItems.Add(draft.Source);
            item.Tag = draft;
            listViewApp.Items.Add(item);
        }
    }

    private void LoadOutlookDrafts()
    {
        if (!_outlookService.IsAvailable) return;

        listViewOutlook.Items.Clear();
        _outlookDrafts = _outlookService.GetDraftsFromOutlook();

        foreach (var draft in _outlookDrafts)
        {
            var item = new ListViewItem(draft.Subject);
            item.SubItems.Add(draft.CreatedAt.ToLocalTime().ToString("g"));
            item.SubItems.Add(draft.UpdatedAt.ToLocalTime().ToString("g"));
            item.SubItems.Add(draft.Source);
            item.Tag = draft;
            listViewOutlook.Items.Add(item);
        }
    }

    private void BtnCreateNew_Click(object? sender, EventArgs e)
    {
        var editorForm = new DraftEditorForm(_dbService, _outlookService, null);
        if (editorForm.ShowDialog() == DialogResult.OK)
        {
            LoadDrafts();
        }
    }

    private void BtnEdit_Click(object? sender, EventArgs e)
    {
        ListView activeList = tabControl.SelectedTab == tabAppDrafts ? listViewApp : listViewOutlook;

        if (activeList.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select a draft to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var draft = activeList.SelectedItems[0].Tag as Draft;
        if (draft != null)
        {
            var editorForm = new DraftEditorForm(_dbService, _outlookService, draft);
            if (editorForm.ShowDialog() == DialogResult.OK)
            {
                LoadDrafts();
            }
        }
    }

    private async void BtnDelete_Click(object? sender, EventArgs e)
    {
        if (listViewApp.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select a draft to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var result = MessageBox.Show("Are you sure you want to delete this draft?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            var draft = listViewApp.SelectedItems[0].Tag as Draft;
            if (draft != null && !string.IsNullOrEmpty(draft.Id))
            {
                await _dbService.Drafts.DeleteOneAsync(d => d.Id == draft.Id);
                LoadDrafts();
            }
        }
    }

    private void BtnPreview_Click(object? sender, EventArgs e)
    {
        ListView activeList = tabControl.SelectedTab == tabAppDrafts ? listViewApp : listViewOutlook;

        if (activeList.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select a draft to preview.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var draft = activeList.SelectedItems[0].Tag as Draft;
        if (draft != null)
        {
            var previewForm = new DraftPreviewForm(draft);
            previewForm.ShowDialog();
        }
    }
}
