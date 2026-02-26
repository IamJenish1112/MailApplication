using MailApplication.Models;
using System.Runtime.InteropServices;

namespace MailApplication.Services;

public class OutlookService
{
    private object? _outlookApp;
    private object? _nameSpace;
    public bool IsAvailable { get; private set; }

    public OutlookService()
    {
        try
        {
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

    /// <summary>
    /// Fetches all configured Outlook accounts with display name, SMTP address, and account type.
    /// </summary>
    public List<EmailAccount> GetOutlookAccounts()
    {
        var accounts = new List<EmailAccount>();
        if (!IsAvailable || _nameSpace == null) return accounts;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var accountsCollection = namespaceType.InvokeMember("Accounts",
                System.Reflection.BindingFlags.GetProperty,
                null, _nameSpace, null);

            if (accountsCollection != null)
            {
                var accountsType = accountsCollection.GetType();
                var count = (int)accountsType.InvokeMember("Count",
                    System.Reflection.BindingFlags.GetProperty,
                    null, accountsCollection, null);

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var account = accountsType.InvokeMember("Item",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null, accountsCollection, new object[] { i });

                        if (account != null)
                        {
                            var accountType = account.GetType();
                            var displayName = accountType.InvokeMember("DisplayName",
                                System.Reflection.BindingFlags.GetProperty,
                                null, account, null) as string ?? "";
                            var smtpAddress = accountType.InvokeMember("SmtpAddress",
                                System.Reflection.BindingFlags.GetProperty,
                                null, account, null) as string ?? "";
                            var acctTypeEnum = accountType.InvokeMember("AccountType",
                                System.Reflection.BindingFlags.GetProperty,
                                null, account, null);

                            string acctTypeName = acctTypeEnum?.ToString() ?? "Unknown";
                            // Map OlAccountType enum values
                            if (int.TryParse(acctTypeName, out int acctTypeVal))
                            {
                                acctTypeName = acctTypeVal switch
                                {
                                    0 => "Exchange",
                                    1 => "IMAP",
                                    2 => "POP3",
                                    3 => "HTTP",
                                    4 => "EAS (ActiveSync)",
                                    5 => "Other",
                                    _ => $"Type {acctTypeVal}"
                                };
                            }

                            accounts.Add(new EmailAccount
                            {
                                AccountName = displayName,
                                EmailAddress = smtpAddress,
                                SmtpAddress = smtpAddress,
                                AccountType = acctTypeName,
                                IsDefault = i == 1 // First account is typically default
                            });

                            Marshal.ReleaseComObject(account);
                        }
                    }
                    catch { /* Skip invalid accounts */ }
                }

                Marshal.ReleaseComObject(accountsCollection);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error fetching Outlook accounts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return accounts;
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

    /// <summary>
    /// Opens the Outlook compose window for designing a template. Returns the HTML body when user saves and closes.
    /// </summary>
    public (string subject, string htmlBody)? OpenOutlookComposer(string? initialSubject = null, string? initialHtmlBody = null)
    {
        if (!IsAvailable || _outlookApp == null) return null;

        try
        {
            var appType = _outlookApp.GetType();
            var mailItem = appType.InvokeMember("CreateItem",
                System.Reflection.BindingFlags.InvokeMethod,
                null, _outlookApp, new object[] { 0 }); // 0 = olMailItem

            if (mailItem != null)
            {
                var itemType = mailItem.GetType();

                if (!string.IsNullOrEmpty(initialSubject))
                    itemType.InvokeMember("Subject", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { initialSubject });

                if (!string.IsNullOrEmpty(initialHtmlBody))
                    itemType.InvokeMember("HTMLBody", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { initialHtmlBody });

                // Display the compose window modally (true = modal)
                itemType.InvokeMember("Display", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, new object[] { true });

                // After the window is closed, capture the content
                var subject = itemType.InvokeMember("Subject",
                    System.Reflection.BindingFlags.GetProperty,
                    null, mailItem, null) as string ?? "";
                var htmlBody = itemType.InvokeMember("HTMLBody",
                    System.Reflection.BindingFlags.GetProperty,
                    null, mailItem, null) as string ?? "";

                // Save it as a draft so we don't lose it
                itemType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, null);

                // Delete the temporary draft from Outlook
                itemType.InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, null);

                Marshal.ReleaseComObject(mailItem);

                return (subject, htmlBody);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening Outlook composer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return null;
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

    /// <summary>
    /// Sends email using BCC only (no TO field). Uses specified account SMTP address if provided.
    /// </summary>
    public void SendEmail(string bcc, string subject, string body, bool isHtml, string? senderSmtpAddress = null)
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

                // Set subject
                itemType.InvokeMember("Subject", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { subject });

                // Set body (HTML or plain text)
                if (isHtml)
                    itemType.InvokeMember("HTMLBody", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });
                else
                    itemType.InvokeMember("Body", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { body });

                // BCC only - no TO field
                if (!string.IsNullOrEmpty(bcc))
                    itemType.InvokeMember("BCC", System.Reflection.BindingFlags.SetProperty, null, mailItem, new object[] { bcc });

                // Set sender account if specified
                if (!string.IsNullOrEmpty(senderSmtpAddress))
                {
                    SetSendUsingAccount(mailItem, senderSmtpAddress);
                }

                itemType.InvokeMember("Send", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, null);
                Marshal.ReleaseComObject(mailItem);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to send email: {ex.Message}");
        }
    }

    /// <summary>
    /// Sets the SendUsingAccount property on a mail item to use a specific Outlook account.
    /// </summary>
    private void SetSendUsingAccount(object mailItem, string smtpAddress)
    {
        if (_nameSpace == null) return;

        try
        {
            var namespaceType = _nameSpace.GetType();
            var accountsCollection = namespaceType.InvokeMember("Accounts",
                System.Reflection.BindingFlags.GetProperty,
                null, _nameSpace, null);

            if (accountsCollection != null)
            {
                var accountsType = accountsCollection.GetType();
                var count = (int)accountsType.InvokeMember("Count",
                    System.Reflection.BindingFlags.GetProperty,
                    null, accountsCollection, null);

                for (int i = 1; i <= count; i++)
                {
                    var account = accountsType.InvokeMember("Item",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null, accountsCollection, new object[] { i });

                    if (account != null)
                    {
                        var accountType = account.GetType();
                        var acctSmtp = accountType.InvokeMember("SmtpAddress",
                            System.Reflection.BindingFlags.GetProperty,
                            null, account, null) as string ?? "";

                        if (acctSmtp.Equals(smtpAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            var itemType = mailItem.GetType();
                            itemType.InvokeMember("SendUsingAccount",
                                System.Reflection.BindingFlags.SetProperty,
                                null, mailItem, new object[] { account });
                            break;
                        }

                        Marshal.ReleaseComObject(account);
                    }
                }

                Marshal.ReleaseComObject(accountsCollection);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not set sender account: {ex.Message}");
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
