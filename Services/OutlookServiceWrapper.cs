using MailApplication.Models;
using System.Runtime.InteropServices;

namespace MailApplication.Services;

public class OutlookService
{
    private object? _outlookApp;
    private object? _nameSpace;
    private Type? _applicationClass;
    private Type? _namespaceClass;
    private Type? _mailItemClass;
    public bool IsAvailable { get; private set; }

    public OutlookService()
    {
        try
        {
            // Try to create Outlook application using COM
            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType != null)
            {
                _outlookApp = Activator.CreateInstance(outlookType);
                if (_outlookApp != null)
                {
                    _nameSpace = outlookType.InvokeMember("GetNamespace",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null, _outlookApp, new object[] { "MAPI" });
                    IsAvailable = true;
                }
            }
        }
        catch (Exception ex)
        {
            IsAvailable = false;
            Console.WriteLine($"Outlook not available: {ex.Message}");
        }
    }

    public List<Draft> GetDraftsFromOutlook()
    {
        var drafts = new List<Draft>();
        if (!IsAvailable || _nameSpace == null) return drafts;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var draftsFolder = namespaceType.InvokeMember("GetDefaultFolder",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _nameSpace, new object[] { 16 }); // 16 = olFolderDrafts

            if (draftsFolder != null)
            {
                var folderType = draftsFolder.GetType();
                var items = folderType.InvokeMember("Items",
                    System.Reflection.BindingFlags.GetProperty,
                    null, draftsFolder, null);

                if (items != null)
                {
                    var itemsType = items.GetType();
                    var count = (int)itemsType.InvokeMember("Count",
                        System.Reflection.BindingFlags.GetProperty,
                        null, items, null);

                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var item = itemsType.InvokeMember("Item",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, items, new object[] { i });

                            if (item != null)
                            {
                                var itemType = item.GetType();
                                var subject = itemType.InvokeMember("Subject",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null) as string ?? "";
                                var htmlBody = itemType.InvokeMember("HTMLBody",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null) as string ?? "";
                                var body = itemType.InvokeMember("Body",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null) as string ?? "";
                                var entryId = itemType.InvokeMember("EntryID",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null) as string ?? "";
                                var creationTime = (DateTime)itemType.InvokeMember("CreationTime",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null);
                                var lastModTime = (DateTime)itemType.InvokeMember("LastModificationTime",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, item, null);

                                drafts.Add(new Draft
                                {
                                    Subject = subject,
                                    Body = htmlBody ?? body,
                                    IsHtml = !string.IsNullOrEmpty(htmlBody),
                                    Source = "Outlook",
                                    OutlookEntryId = entryId,
                                    CreatedAt = creationTime,
                                    UpdatedAt = lastModTime
                                });

                                Marshal.ReleaseComObject(item);
                            }
                        }
                        catch { /* Skip invalid items */ }
                    }

                    Marshal.ReleaseComObject(items);
                }
                Marshal.ReleaseComObject(draftsFolder);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading Outlook drafts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return drafts;
    }

    public void OpenDraftInOutlook(string entryId)
    {
        if (!IsAvailable || _nameSpace == null || string.IsNullOrEmpty(entryId)) return;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var mailItem = namespaceType.InvokeMember("GetItemFromID",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _nameSpace, new object[] { entryId });

            if (mailItem != null)
            {
                var itemType = mailItem.GetType();
                itemType.InvokeMember("Display",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, mailItem, new object[] { false });
                Marshal.ReleaseComObject(mailItem);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening draft: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public List<object> GetInboxItems()
    {
        var items = new List<object>();
        if (!IsAvailable || _nameSpace == null) return items;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var inboxFolder = namespaceType.InvokeMember("GetDefaultFolder",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _nameSpace, new object[] { 6 }); // 6 = olFolderInbox

            if (inboxFolder != null)
            {
                var folderType = inboxFolder.GetType();
                var folderItems = folderType.InvokeMember("Items",
                    System.Reflection.BindingFlags.GetProperty,
                    null, inboxFolder, null);

                if (folderItems != null)
                {
                    var itemsType = folderItems.GetType();
                    var count = Math.Min(100, (int)itemsType.InvokeMember("Count",
                        System.Reflection.BindingFlags.GetProperty,
                        null, folderItems, null));

                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var item = itemsType.InvokeMember("Item",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, folderItems, new object[] { i });
                            if (item != null) items.Add(item);
                        }
                        catch { /* Skip invalid items */ }
                    }

                    Marshal.ReleaseComObject(folderItems);
                }
                Marshal.ReleaseComObject(inboxFolder);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading inbox: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return items;
    }

    public List<object> GetSentItems()
    {
        var items = new List<object>();
        if (!IsAvailable || _nameSpace == null) return items;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var sentFolder = namespaceType.InvokeMember("GetDefaultFolder",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _nameSpace, new object[] { 5 }); // 5 = olFolderSentMail

            if (sentFolder != null)
            {
                var folderType = sentFolder.GetType();
                var folderItems = folderType.InvokeMember("Items",
                    System.Reflection.BindingFlags.GetProperty,
                    null, sentFolder, null);

                if (folderItems != null)
                {
                    var itemsType = folderItems.GetType();
                    var count = Math.Min(100, (int)itemsType.InvokeMember("Count",
                        System.Reflection.BindingFlags.GetProperty,
                        null, folderItems, null));

                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var item = itemsType.InvokeMember("Item",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, folderItems, new object[] { i });
                            if (item != null) items.Add(item);
                        }
                        catch { /* Skip invalid items */ }
                    }

                    Marshal.ReleaseComObject(folderItems);
                }
                Marshal.ReleaseComObject(sentFolder);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading sent items: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return items;
    }

    public void SendEmail(string bcc, string subject, string body, bool isHtml)
    {
        if (!IsAvailable || _outlookApp == null) return;

        try
        {
            var appType = _outlookApp.GetType();
            var mailItem = appType.InvokeMember("CreateItem",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _outlookApp, new object[] { 0 }); // 0 = olMailItem

            if (mailItem != null)
            {
                var itemType = mailItem.GetType();
                // itemType.InvokeMember("To", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { "" });
                itemType.InvokeMember("Subject", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { subject });

                if (isHtml)
                    itemType.InvokeMember("HTMLBody", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });
                else
                    itemType.InvokeMember("Body", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });

                if (!string.IsNullOrEmpty(bcc))
                    itemType.InvokeMember("BCC", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { bcc });

                itemType.InvokeMember("Send", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, null);
                Marshal.ReleaseComObject(mailItem);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to send email: {ex.Message}");
        }
    }

    public void SaveDraftToOutlook(string subject, string body, bool isHtml)
    {
        if (!IsAvailable || _outlookApp == null) return;

        try
        {
            var appType = _outlookApp.GetType();
            var mailItem = appType.InvokeMember("CreateItem",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _outlookApp, new object[] { 0 }); // 0 = olMailItem

            if (mailItem != null)
            {
                var itemType = mailItem.GetType();
                itemType.InvokeMember("Subject", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { subject });

                if (isHtml)
                    itemType.InvokeMember("HTMLBody", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });
                else
                    itemType.InvokeMember("Body", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });

                itemType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, null);
                Marshal.ReleaseComObject(mailItem);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to save draft to Outlook: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_nameSpace != null)
        {
            Marshal.ReleaseComObject(_nameSpace);
            _nameSpace = null;
        }

        if (_outlookApp != null)
        {
            Marshal.ReleaseComObject(_outlookApp);
            _outlookApp = null;
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
