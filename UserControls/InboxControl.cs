using MailApplication.Services;

namespace MailApplication.UserControls;

public partial class InboxControl : UserControl
{
    private readonly OutlookService _outlookService;
    private ListView listViewInbox;
    private Button btnRefresh;
    private Button btnOpen;

    public InboxControl(OutlookService outlookService)
    {
        _outlookService = outlookService;
        InitializeComponent();
        SetupUI();

        if (!_outlookService.IsAvailable)
        {
            var label = new Label
            {
                Text = "Outlook is not available on this system.\nInbox features are disabled.",
                Font = new Font("Segoe UI", 12),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
            this.Controls.Add(label);
            btnRefresh.Enabled = false;
            btnOpen.Enabled = false;
        }
        else
        {
            LoadInbox();
        }
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);

        var titleLabel = new Label
        {
            Text = "Inbox",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        listViewInbox = new ListView
        {
            Location = new Point(0, 60),
            Size = new Size(980, 500),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewInbox.Columns.Add("From", 250);
        listViewInbox.Columns.Add("Subject", 400);
        listViewInbox.Columns.Add("Received", 180);
        listViewInbox.Columns.Add("Size", 150);

        var buttonPanel = new Panel
        {
            Location = new Point(0, 580),
            Size = new Size(980, 50),
            BackColor = Color.White
        };

        btnRefresh = new Button
        {
            Text = "Refresh",
            Location = new Point(0, 5),
            Size = new Size(120, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnRefresh.Click += (s, e) => LoadInbox();

        btnOpen = new Button
        {
            Text = "Open Selected",
            Location = new Point(140, 5),
            Size = new Size(140, 40),
            BackColor = Color.FromArgb(13, 110, 253),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnOpen.Click += BtnOpen_Click;

        buttonPanel.Controls.AddRange(new Control[] { btnRefresh, btnOpen });

        this.Controls.AddRange(new Control[] { titleLabel, listViewInbox, buttonPanel });
    }

    private void LoadInbox()
    {
        listViewInbox.Items.Clear();
        var items = _outlookService.GetInboxItems();

        foreach (var mailItem in items.Take(100))
        {
            try
            {
                var itemType = mailItem.GetType();
                var senderName = itemType.InvokeMember("SenderName", System.Reflection.BindingFlags.GetProperty, null, mailItem, null) as string ?? "Unknown";
                var subject = itemType.InvokeMember("Subject", System.Reflection.BindingFlags.GetProperty, null, mailItem, null) as string ?? "(No Subject)";
                var receivedTime = (DateTime)itemType.InvokeMember("ReceivedTime", System.Reflection.BindingFlags.GetProperty, null, mailItem, null);
                var size = (int)itemType.InvokeMember("Size", System.Reflection.BindingFlags.GetProperty, null, mailItem, null);

                var item = new ListViewItem(senderName);
                item.SubItems.Add(subject);
                item.SubItems.Add(receivedTime.ToString("g"));
                item.SubItems.Add($"{size / 1024} KB");
                item.Tag = mailItem;
                listViewInbox.Items.Add(item);
            }
            catch { /* Skip invalid items */ }
        }
    }

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        if (listViewInbox.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an email to open.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var mailItem = listViewInbox.SelectedItems[0].Tag;
        if (mailItem != null)
        {
            var itemType = mailItem.GetType();
            itemType.InvokeMember("Display", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, new object[] { false });
        }
    }
}
