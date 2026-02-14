using MailApplication.Services;

namespace MailApplication.UserControls;

public partial class OutboxControl : UserControl
{
    private readonly OutlookService _outlookService;
    private ListView listViewOutbox;
    private Button btnRefresh;
    private Button btnOpen;

    public OutboxControl(OutlookService outlookService)
    {
        _outlookService = outlookService;
        InitializeComponent();
        SetupUI();

        if (!_outlookService.IsAvailable)
        {
            var label = new Label
            {
                Text = "Outlook is not available on this system.\nOutbox features are disabled.",
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
            LoadOutbox();
        }
    }

    private void SetupUI()
    {
        this.BackColor = Color.White;
        this.Padding = new Padding(20);

        var titleLabel = new Label
        {
            Text = "Sent Items",
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41),
            AutoSize = true,
            Location = new Point(0, 0)
        };

        listViewOutbox = new ListView
        {
            Location = new Point(0, 60),
            Size = new Size(980, 500),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            Font = new Font("Segoe UI", 10)
        };

        listViewOutbox.Columns.Add("To", 200);
        listViewOutbox.Columns.Add("Subject", 300);
        listViewOutbox.Columns.Add("From Account", 200);
        listViewOutbox.Columns.Add("Sent", 150);
        listViewOutbox.Columns.Add("Size", 130);

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
        btnRefresh.Click += (s, e) => LoadOutbox();

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

        this.Controls.AddRange(new Control[] { titleLabel, listViewOutbox, buttonPanel });
    }

    private void LoadOutbox()
    {
        listViewOutbox.Items.Clear();
        var items = _outlookService.GetSentItems();

        foreach (var mailItem in items.Take(100))
        {
            try
            {
                var itemType = mailItem.GetType();
                var to = itemType.InvokeMember("To", System.Reflection.BindingFlags.GetProperty, null, mailItem, null) as string ?? "Unknown";
                var subject = itemType.InvokeMember("Subject", System.Reflection.BindingFlags.GetProperty, null, mailItem, null) as string ?? "(No Subject)";

                var senderObj = itemType.InvokeMember("Sender", System.Reflection.BindingFlags.GetProperty, null, mailItem, null);
                var fromAccount = "Unknown";
                if (senderObj != null)
                {
                    var senderType = senderObj.GetType();
                    fromAccount = senderType.InvokeMember("Address", System.Reflection.BindingFlags.GetProperty, null, senderObj, null) as string ?? "Unknown";
                }

                var sentOn = (DateTime)itemType.InvokeMember("SentOn", System.Reflection.BindingFlags.GetProperty, null, mailItem, null);
                var size = (int)itemType.InvokeMember("Size", System.Reflection.BindingFlags.GetProperty, null, mailItem, null);

                var item = new ListViewItem(to);
                item.SubItems.Add(subject);
                item.SubItems.Add(fromAccount);
                item.SubItems.Add(sentOn.ToString("g"));
                item.SubItems.Add($"{size / 1024} KB");
                item.Tag = mailItem;
                listViewOutbox.Items.Add(item);
            }
            catch { /* Skip invalid items */ }
        }
    }

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        if (listViewOutbox.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select an email to open.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var mailItem = listViewOutbox.SelectedItems[0].Tag;
        if (mailItem != null)
        {
            var itemType = mailItem.GetType();
            itemType.InvokeMember("Display", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, new object[] { false });
        }
    }
}
