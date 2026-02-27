using MailApplication.Models;
using System.Runtime.InteropServices;

namespace MailApplication.Services;

/// <summary>
/// Service responsible for managing Outlook accounts and sending emails through them.
/// 
/// KEY FIX: COM account objects are cached at initialization and kept alive for the
/// entire service lifetime. This prevents GC from reclaiming the COM RCW before
/// SendUsingAccount can take effect. Previously, account COM objects were fetched
/// and released per-send, causing Outlook to always fall back to the default account.
/// </summary>
public class OutlookAccountService : IDisposable
{
    private object? _outlookApp;
    private object? _nameSpace;
    private readonly object _comLock = new();
    private bool _disposed = false;

    // CRITICAL: Cache raw COM account objects so they stay alive for SendUsingAccount
    private readonly Dictionary<string, object> _cachedComAccounts = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<EmailAccount> _cachedAccountModels = new();

    private static readonly string LogFile = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "outlook_account_service.log");

    public bool IsAvailable { get; private set; }

    /// <summary>
    /// Number of accounts currently cached and available for sending.
    /// </summary>
    public int AccountCount => _cachedAccountModels.Count;

    // P/Invoke for GetActiveObject (.NET 5+ removed Marshal.GetActiveObject)
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid clsid,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public OutlookAccountService()
    {
        InitializeOutlook();
        if (IsAvailable)
        {
            CacheAllAccounts();
        }
    }

    /// <summary>
    /// Initializes the Outlook COM application instance.
    /// Re-uses existing Outlook instance if already running.
    /// </summary>
    private void InitializeOutlook()
    {
        try
        {
            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");

            // Try to connect to existing Outlook instance first (better performance)
            try
            {
                if (outlookType != null)
                {
                    GetActiveObject(outlookType.GUID, IntPtr.Zero, out object activeObj);
                    _outlookApp = activeObj;
                    Log("Connected to existing Outlook instance.");
                }
            }
            catch (COMException)
            {
                // Outlook not running, create new instance
                if (outlookType != null)
                {
                    _outlookApp = Activator.CreateInstance(outlookType);
                    Log("Created new Outlook instance.");
                }
            }

            if (_outlookApp != null)
            {
                _nameSpace = _outlookApp.GetType().InvokeMember("GetNamespace",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, _outlookApp, new object[] { "MAPI" });
                IsAvailable = true;
            }
        }
        catch (Exception ex)
        {
            IsAvailable = false;
            Log($"ERROR: Outlook not available: {ex.Message}");
        }
    }

    /// <summary>
    /// Caches all Outlook account COM objects at startup.
    /// These COM references stay alive for the entire service lifetime,
    /// which is REQUIRED for SendUsingAccount to work correctly.
    /// </summary>
    private void CacheAllAccounts()
    {
        if (_nameSpace == null) return;

        object? accountsCollection = null;

        try
        {
            lock (_comLock)
            {
                var namespaceType = _nameSpace.GetType();
                accountsCollection = namespaceType.InvokeMember("Accounts",
                    System.Reflection.BindingFlags.GetProperty,
                    null, _nameSpace, null);

                if (accountsCollection == null) return;

                var accountsType = accountsCollection.GetType();
                var count = (int)accountsType.InvokeMember("Count",
                    System.Reflection.BindingFlags.GetProperty,
                    null, accountsCollection, null);

                Log($"Caching {count} Outlook account(s)...");

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var account = accountsType.InvokeMember("Item",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null, accountsCollection, new object[] { i });

                        if (account != null)
                        {
                            var displayName = SafeGetProperty<string>(account, "DisplayName") ?? "";
                            var smtpAddress = SafeGetProperty<string>(account, "SmtpAddress") ?? "";
                            var acctTypeEnum = SafeGetProperty<object>(account, "AccountType");
                            string acctTypeName = MapAccountType(acctTypeEnum);

                            if (!string.IsNullOrEmpty(smtpAddress))
                            {
                                // KEEP the COM reference alive — DO NOT release it
                                _cachedComAccounts[smtpAddress] = account;

                                _cachedAccountModels.Add(new EmailAccount
                                {
                                    AccountName = displayName,
                                    EmailAddress = smtpAddress,
                                    SmtpAddress = smtpAddress,
                                    AccountType = acctTypeName,
                                    IsDefault = i == 1,
                                    IsEnabled = true
                                });

                                Log($"  ✓ Cached Account {i}: {displayName} ({smtpAddress}) - {acctTypeName}");
                            }
                            else
                            {
                                SafeReleaseCom(account);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"  WARNING: Skipping account {i}: {ex.Message}");
                    }
                }

                // DO NOT release accountsCollection — keep it alive too
                // (releasing it can invalidate child account references in some Outlook versions)
            }
        }
        catch (Exception ex)
        {
            Log($"ERROR: Failed to cache accounts: {ex.Message}");
        }
    }

    /// <summary>
    /// Returns the list of all cached Outlook accounts.
    /// These are model objects (not COM), safe to use anywhere.
    /// </summary>
    public List<EmailAccount> GetAllAccounts()
    {
        return _cachedAccountModels.ToList();
    }

    /// <summary>
    /// Sends a single email using the specified Outlook account.
    /// Uses the CACHED COM account object for SendUsingAccount (reliable).
    /// </summary>
    public bool SendEmail(string bccRecipients, string subject, string body, bool isHtml, string senderSmtpAddress, string? replyToEmail = null)
    {
        if (!IsAvailable || _outlookApp == null)
            throw new InvalidOperationException("Outlook is not available.");

        object? mailItem = null;

        try
        {
            lock (_comLock)
            {
                var appType = _outlookApp.GetType();
                mailItem = appType.InvokeMember("CreateItem",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, _outlookApp, new object[] { 0 }); // 0 = olMailItem

                if (mailItem == null)
                    throw new InvalidOperationException("Failed to create mail item.");

                var itemType = mailItem.GetType();

                // 1. Set SendUsingAccount FIRST using the cached COM object
                if (_cachedComAccounts.TryGetValue(senderSmtpAddress, out var cachedAccount))
                {
                    itemType.InvokeMember("SendUsingAccount",
                        System.Reflection.BindingFlags.SetProperty,
                        null, mailItem, new object[] { cachedAccount });
                    Log($"SendUsingAccount set to: {senderSmtpAddress}");
                }
                else
                {
                    Log($"WARNING: No cached account for '{senderSmtpAddress}'. Will use default account.");
                }

                // 2. Set subject
                itemType.InvokeMember("Subject",
                    System.Reflection.BindingFlags.SetProperty,
                    null, mailItem, new object[] { subject });

                // 3. Set body (HTML or plain text)
                if (isHtml)
                {
                    itemType.InvokeMember("HTMLBody",
                        System.Reflection.BindingFlags.SetProperty,
                        null, mailItem, new object[] { body });
                }
                else
                {
                    itemType.InvokeMember("Body",
                        System.Reflection.BindingFlags.SetProperty,
                        null, mailItem, new object[] { body });
                }

                // 4. BCC only — no TO field
                if (!string.IsNullOrEmpty(bccRecipients))
                {
                    itemType.InvokeMember("BCC",
                        System.Reflection.BindingFlags.SetProperty,
                        null, mailItem, new object[] { bccRecipients });
                }

                // 5. Set Reply-To header if specified
                if (!string.IsNullOrWhiteSpace(replyToEmail))
                {
                    try
                    {
                        // Use Outlook PropertyAccessor to set the Reply-To Internet header
                        // PR_REPLY_RECIPIENT_NAMES is the simplest approach via headers
                        const string PR_REPLY_RECIPIENT_NAMES = "http://schemas.microsoft.com/mapi/proptag/0x0050001E";
                        var propAccessor = itemType.InvokeMember("PropertyAccessor",
                            System.Reflection.BindingFlags.GetProperty,
                            null, mailItem, null);
                        if (propAccessor != null)
                        {
                            propAccessor.GetType().InvokeMember("SetProperty",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, propAccessor, new object[] { PR_REPLY_RECIPIENT_NAMES, replyToEmail });
                            Log($"Reply-To set to: {replyToEmail}");
                        }
                    }
                    catch (Exception rtEx)
                    {
                        Log($"WARNING: Could not set Reply-To header: {rtEx.Message}");
                    }
                }

                // 6. Send
                itemType.InvokeMember("Send",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, mailItem, null);

                Log($"✓ Email SENT via [{senderSmtpAddress}] to {bccRecipients.Split(';').Length} recipient(s). ReplyTo={replyToEmail ?? "(none)"}");
                return true;
            }
        }
        catch (Exception ex)
        {
            Log($"ERROR: Failed to send email via {senderSmtpAddress}: {ex.Message}");
            throw;
        }
        finally
        {
            SafeReleaseCom(mailItem);
        }
    }

    #region Helper Methods

    private static string MapAccountType(object? acctTypeEnum)
    {
        string acctTypeName = acctTypeEnum?.ToString() ?? "Unknown";
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
        return acctTypeName;
    }

    private static T? SafeGetProperty<T>(object comObject, string propertyName)
    {
        try
        {
            var result = comObject.GetType().InvokeMember(propertyName,
                System.Reflection.BindingFlags.GetProperty,
                null, comObject, null);
            return result is T typed ? typed : default;
        }
        catch
        {
            return default;
        }
    }

    private static void SafeReleaseCom(object? comObject)
    {
        if (comObject != null)
        {
            try { Marshal.ReleaseComObject(comObject); }
            catch { /* already released */ }
        }
    }

    private static void Log(string message)
    {
        try
        {
            var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [OutlookAccountService] {message}";
            File.AppendAllText(LogFile, logEntry + Environment.NewLine);
            Console.WriteLine(logEntry);
        }
        catch { }
    }

    #endregion

    #region IDisposable

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        // Release all cached COM account objects
        foreach (var kvp in _cachedComAccounts)
        {
            SafeReleaseCom(kvp.Value);
        }
        _cachedComAccounts.Clear();
        _cachedAccountModels.Clear();

        SafeReleaseCom(_nameSpace);
        _nameSpace = null;

        SafeReleaseCom(_outlookApp);
        _outlookApp = null;

        GC.Collect();
        GC.WaitForPendingFinalizers();

        _disposed = true;
        Log("OutlookAccountService disposed. All COM objects released.");
    }

    ~OutlookAccountService()
    {
        Dispose(false);
    }

    #endregion
}
