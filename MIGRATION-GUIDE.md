# Migration Guide: From MailKit to CloudMailKit

This guide shows you how to migrate your existing MailKit/MimeKit code to CloudMailKit for Microsoft Graph OAuth support.

## Why Migrate?

- **Microsoft 365** is deprecating basic authentication and SMTP
- **OAuth 2.0** is now required for many email services
- **Cloud-first** approach with Graph API
- **Same code, different backend** - your application logic stays the same

## The Good News: It's a Drop-In Replacement!

CloudMailKit provides **100% API compatibility** with MailKit/MimeKit. You only need to change:

1. ✅ The `using` statement (namespace)
2. ✅ The authentication credentials (OAuth instead of SMTP password)

Everything else stays **exactly the same**!

## Step-by-Step Migration

### Step 1: Change the Using Statement

**Before:**
```csharp
using MailKit.Net.Smtp;
using MimeKit;
```

**After:**
```csharp
using CloudMailKit.MailKit;
```

### Step 2: Update Authentication

**Before (SMTP with basic auth):**
```csharp
using (var client = new SmtpClient())
{
    client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
    client.Authenticate("user@company.com", "password");
    client.Send(message);
    client.Disconnect(true);
}
```

**After (Graph API with OAuth):**
```csharp
using (var client = new SmtpClient())
{
    client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
    client.Authenticate("client-id|tenant-id|user@company.com", "client-secret");
    client.Send(message);
    client.Disconnect(true);
}
```

### Step 3: Done! 🎉

That's it! Your entire application continues to work with **zero changes** to:
- Message creation
- Body building
- Attachment handling
- Address management
- All other MailKit features

## Complete Example: Before and After

### BEFORE (MailKit/MimeKit with SMTP)

```csharp
using MailKit.Net.Smtp;
using MimeKit;

public void SendEmail()
{
    var message = new MimeMessage();
    message.From.Add(new MailboxAddress("Sales", "sales@company.com"));
    message.To.Add(new MailboxAddress("John Doe", "john@example.com"));
    message.Cc.Add(new MailboxAddress("Manager", "manager@company.com"));
    message.Subject = "Monthly Report";
    message.Importance = MessageImportance.High;

    var builder = new BodyBuilder();
    builder.TextBody = "Please find the report attached.";
    builder.HtmlBody = "<p>Please find the report <strong>attached</strong>.</p>";
    builder.AddAttachment("report.pdf");
    message.Body = builder.ToMessageBody();

    using (var client = new SmtpClient())
    {
        client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
        client.Authenticate("sales@company.com", "password123");
        client.Send(message);
        client.Disconnect(true);
    }
}
```

### AFTER (CloudMailKit with Graph OAuth)

```csharp
using CloudMailKit.MailKit;  // <-- ONLY CHANGE #1

public void SendEmail()
{
    var message = new MimeMessage();
    message.From.Add(new MailboxAddress("Sales", "sales@company.com"));
    message.To.Add(new MailboxAddress("John Doe", "john@example.com"));
    message.Cc.Add(new MailboxAddress("Manager", "manager@company.com"));
    message.Subject = "Monthly Report";
    message.Importance = MessageImportance.High;

    var builder = new BodyBuilder();
    builder.TextBody = "Please find the report attached.";
    builder.HtmlBody = "<p>Please find the report <strong>attached</strong>.</p>";
    builder.AddAttachment("report.pdf");
    message.Body = builder.ToMessageBody();

    using (var client = new SmtpClient())
    {
        client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
        // ONLY CHANGE #2: OAuth credentials instead of username/password
        client.Authenticate("app-id|tenant-id|sales@company.com", "client-secret");
        client.Send(message);
        client.Disconnect(true);
    }
}
```

## What Stays the Same?

Everything! Including:

✅ **MimeMessage** - Create and manipulate messages
✅ **MailboxAddress** - Email addresses with names
✅ **BodyBuilder** - Build text/HTML bodies
✅ **TextPart** - Plain text and HTML parts
✅ **Multipart** - Multipart messages
✅ **MimePart** - File attachments
✅ **InternetAddressList** - Manage recipient lists
✅ **SmtpClient** - Send emails (now via Graph API)
✅ **Async support** - All `*Async()` methods work
✅ **Attachments** - File attachments work identically
✅ **Importance/Priority** - Message flags
✅ **Multiple recipients** - To, Cc, Bcc lists

## Azure AD Setup

To get your OAuth credentials:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Name your app and register it
5. Copy the **Application (client) ID** and **Directory (tenant) ID**
6. Go to **Certificates & secrets** → Create a **New client secret**
7. Copy the secret value immediately (it won't be shown again)
8. Go to **API permissions** → **Add a permission**
9. Select **Microsoft Graph** → **Application permissions**
10. Add **Mail.Send** permission
11. Click **Grant admin consent**

Now use these values:
- `client-id` = Application (client) ID
- `tenant-id` = Directory (tenant) ID
- `client-secret` = Secret value

## Alternative: Explicit Credentials

If you prefer not to use the pipe-delimited format:

```csharp
using (var client = new SmtpClient())
{
    client.SetGraphCredentials(
        tenantId: "your-tenant-id",
        clientId: "your-client-id",
        clientSecret: "your-client-secret",
        senderAddress: "sender@company.com"
    );
    client.Send(message);
}
```

## Migrating Async Code

**Before:**
```csharp
await client.ConnectAsync("smtp.office365.com", 587, SecureSocketOptions.StartTls);
await client.AuthenticateAsync("user@company.com", "password");
await client.SendAsync(message);
await client.DisconnectAsync(true);
```

**After:**
```csharp
await client.ConnectAsync("graph.microsoft.com", 0, SecureSocketOptions.Auto);
await client.AuthenticateAsync("client-id|tenant-id|user@company.com", "secret");
await client.SendAsync(message);
await client.DisconnectAsync(true);
```

## Configuration File Support

Store credentials in `app.config` or `web.config`:

```xml
<configuration>
  <appSettings>
    <add key="CloudMailKit.TenantId" value="your-tenant-id" />
    <add key="CloudMailKit.ClientId" value="your-client-id" />
    <add key="CloudMailKit.ClientSecret" value="your-client-secret" />
    <add key="CloudMailKit.MailboxAddress" value="sender@company.com" />
  </appSettings>
</configuration>
```

Then use:

```csharp
var client = new UnifiedMailClient();
client.InitializeFromConfig();
client.SendSimple(
    from: "sender@company.com",
    to: "recipient@example.com",
    subject: "Test",
    body: "Hello!",
    isHtml: false
);
```

## Testing Your Migration

1. **Create a test project** with both MailKit and CloudMailKit
2. **Write a test** that sends an email using MailKit
3. **Copy the test** and change only the namespace
4. **Verify** both tests produce identical emails

Example test:

```csharp
[TestMethod]
public void TestMailKitCompatibility()
{
    // This code is IDENTICAL to MailKit
    var message = new MimeMessage();
    message.From.Add(new MailboxAddress("Test", "test@company.com"));
    message.To.Add(new MailboxAddress("Recipient", "recipient@example.com"));
    message.Subject = "Test";
    message.Body = new TextPart("plain") { Text = "Hello!" };

    using (var client = new SmtpClient())
    {
        client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
        client.Authenticate("client-id|tenant-id|test@company.com", "secret");
        client.Send(message);
        client.Disconnect(true);
    }
}
```

## Common Pitfalls

### ❌ Wrong: Still using SMTP host
```csharp
client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
```

### ✅ Right: Use Graph API endpoint
```csharp
client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
```

### ❌ Wrong: Using old credentials format
```csharp
client.Authenticate("user@company.com", "password123");
```

### ✅ Right: Using OAuth credentials
```csharp
client.Authenticate("client-id|tenant-id|user@company.com", "client-secret");
```

## Troubleshooting

### "Not authenticated" error
- Verify your Azure AD app has Mail.Send permission
- Ensure admin consent was granted
- Check client ID, tenant ID, and secret are correct

### "Invalid credentials" error
- Format must be: `client-id|tenant-id|sender@company.com`
- Use pipe character `|` to separate values
- Ensure no extra spaces

### "Send failed" error
- Verify the sender address has a mailbox in your tenant
- Check the sender address matches the authenticated user
- Ensure the app has proper permissions

## Need Help?

- 📚 [Full Documentation](README.md)
- 💬 [GitHub Issues](https://github.com/PSScript/CloudMailKit/issues)
- 📧 [Email Support](mailto:support@cloudmailkit.com)

## Summary

Migrating from MailKit to CloudMailKit is simple:

1. Change `using MailKit` → `using CloudMailKit.MailKit`
2. Change authentication to OAuth format
3. Done!

Your entire codebase continues to work without any other changes. It's a true drop-in replacement! 🚀
