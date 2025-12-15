# CloudMailKit

**🚨 Rescue Your Legacy Applications Before SMTP Dies! 🚨**

A **100% drop-in replacement** for MailKit/MimeKit and legacy SMTP code that uses **Microsoft Graph OAuth** instead of basic authentication.

## 🆘 Is Your App About to Break?

**Microsoft is disabling SMTP with basic auth** - your old applications will STOP sending emails:

- ❌ VB6 apps from the 2000s
- ❌ Classic ASP e-commerce sites
- ❌ .NET Framework 2.0-4.8 applications
- ❌ Magic xpa business systems
- ❌ Windows Services that email alerts
- ❌ Scheduled tasks using CDO/SMTP

**CloudMailKit is your emergency fix** ✅
- Install in **minutes**, not months
- **Minimal code changes** required
- Buy time to **modernize properly**
- Keep business running **NOW**

## Why CloudMailKit?

Modern cloud email services like Microsoft 365 are deprecating basic authentication and SMTP. CloudMailKit provides an identical API to MailKit/MimeKit but uses Microsoft Graph API with OAuth authentication under the hood.

**Perfect for:**
- 🏚️ **Legacy applications** (VB6, Classic ASP, old .NET Framework)
- 🔄 **Migration projects** (MailKit/MimeKit users)
- 🚀 **New development** (.NET 6+ applications)

### Key Features

✅ **100% API Compatible** - Drop-in replacement for MailKit/MimeKit
✅ **Microsoft Graph OAuth** - No SMTP required
✅ **Zero Code Changes** - Same classes, same methods, same syntax
✅ **COM Visible** - Works with VB6, Magic xpa, and other COM clients
✅ **Full Email Support** - Text, HTML, attachments, CC, BCC, importance
✅ **Email Reading** - Fetch, search, and manage emails via Graph API
✅ **.NET Framework 4.8** - Works with stone age applications

---

## 📚 Documentation Quick Links

- 🆘 **[LEGACY RESCUE GUIDE](LEGACY-RESCUE-GUIDE.md)** ← **Start here for VB6/Classic ASP/old .NET**
- 🔄 **[Migration Guide](MIGRATION-GUIDE.md)** ← For MailKit/MimeKit users
- 📖 **[Full Documentation](#)** ← Complete API reference (below)

---

## Installation

```bash
dotnet add package CloudMailKit
```

Or add to your `.csproj`:

```xml
<PackageReference Include="CloudMailKit" Version="1.0.0" />
```

## Quick Start

### Using the MailKit-Compatible API (Recommended)

**Before (Original MailKit code):**

```csharp
using MailKit.Net.Smtp;
using MimeKit;

var message = new MimeMessage();
message.From.Add(new MailboxAddress("Sender Name", "sender@domain.com"));
message.To.Add(new MailboxAddress("Recipient", "recipient@domain.com"));
message.Subject = "Hello from MailKit";

var builder = new BodyBuilder();
builder.TextBody = "This is the plain text body";
builder.HtmlBody = "<p>This is the <strong>HTML</strong> body</p>";
message.Body = builder.ToMessageBody();

using (var client = new SmtpClient())
{
    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
    client.Authenticate("username", "password");
    client.Send(message);
    client.Disconnect(true);
}
```

**After (CloudMailKit - IDENTICAL API):**

```csharp
using CloudMailKit.MailKit;  // <-- Only change: namespace

var message = new MimeMessage();
message.From.Add(new MailboxAddress("Sender Name", "sender@domain.com"));
message.To.Add(new MailboxAddress("Recipient", "recipient@domain.com"));
message.Subject = "Hello from CloudMailKit";

var builder = new BodyBuilder();
builder.TextBody = "This is the plain text body";
builder.HtmlBody = "<p>This is the <strong>HTML</strong> body</p>";
builder.AddAttachment("path/to/file.pdf");
message.Body = builder.ToMessageBody();

using (var client = new SmtpClient())
{
    // Format: "client_id|tenant_id|sender@domain.com"
    client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
    client.Authenticate("your-client-id|your-tenant-id|sender@domain.com", "your-client-secret");
    client.Send(message);
    client.Disconnect(true);
}
```

### Authentication Format

CloudMailKit uses the username parameter to pass OAuth credentials:

```
Username: "client_id|tenant_id|sender@domain.com"
Password: "client_secret"
```

Or use the explicit method:

```csharp
var client = new SmtpClient();
client.SetGraphCredentials(
    tenantId: "your-tenant-id",
    clientId: "your-client-id",
    clientSecret: "your-client-secret",
    senderAddress: "sender@domain.com"
);
client.Send(message);
```

## Migration Guide

### Step 1: Replace the Namespace

```csharp
// Old
using MailKit.Net.Smtp;
using MimeKit;

// New
using CloudMailKit.MailKit;
```

### Step 2: Update Authentication

```csharp
// Old SMTP
client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
client.Authenticate("user@domain.com", "password");

// New Graph OAuth
client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
client.Authenticate("client-id|tenant-id|user@domain.com", "client-secret");
```

### Step 3: That's it! 🎉

Everything else stays exactly the same:
- `MimeMessage` - ✅ Same
- `MailboxAddress` - ✅ Same
- `BodyBuilder` - ✅ Same
- `TextPart` - ✅ Same
- `Multipart` - ✅ Same
- `MimePart` - ✅ Same

## Complete Examples

### Sending HTML Email with Attachments

```csharp
using CloudMailKit.MailKit;

var message = new MimeMessage();
message.From.Add(new MailboxAddress("Sales Team", "sales@company.com"));
message.To.Add(new MailboxAddress("John Doe", "john@example.com"));
message.Cc.Add(new MailboxAddress("Manager", "manager@company.com"));
message.Subject = "Q4 Sales Report";
message.Importance = MessageImportance.High;

var builder = new BodyBuilder();
builder.HtmlBody = @"
    <h1>Q4 Sales Report</h1>
    <p>Please find the attached report for Q4 2024.</p>
    <p>Best regards,<br/>Sales Team</p>
";
builder.AddAttachment("Q4-Report.pdf");
builder.AddAttachment("Summary.xlsx");
message.Body = builder.ToMessageBody();

using (var client = new SmtpClient())
{
    client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
    client.Authenticate("app-id|tenant-id|sales@company.com", "secret");
    await client.SendAsync(message);
    client.Disconnect(true);
}
```

### Sending with Alternative Text/HTML

```csharp
var message = new MimeMessage();
message.From.Add(new MailboxAddress("Newsletter", "news@company.com"));
message.To.Add(new MailboxAddress("Subscriber", "user@example.com"));
message.Subject = "Weekly Newsletter";

var builder = new BodyBuilder();
builder.TextBody = "This is the plain text version of the newsletter.";
builder.HtmlBody = "<h1>Weekly Newsletter</h1><p>This is the <strong>HTML</strong> version.</p>";
message.Body = builder.ToMessageBody();

using (var client = new SmtpClient())
{
    client.SetGraphCredentials("tenant-id", "client-id", "secret", "news@company.com");
    client.Send(message);
}
```

## Advanced Features

### Reading Emails

```csharp
using CloudMailKit;

var client = new UnifiedMailClient();
client.Initialize(
    tenantId: "your-tenant-id",
    clientId: "your-client-id",
    clientSecret: "your-client-secret",
    mailboxAddress: "user@domain.com"
);

// Get inbox
var inbox = client.GetInbox();
Console.WriteLine($"Total messages: {inbox.TotalItemCount}");

// List unread messages
var unread = client.GetUnreadMessages(50);
foreach (var msg in unread)
{
    Console.WriteLine($"{msg.Subject} - {msg.From}");
}

// Search messages
var results = client.SearchMessages("from:important@domain.com");

// Get message with MIME content
var message = client.GetMessage(messageId);
var mimeContent = client.GetMessageMime(messageId);

// Parse MIME content
var subject = client.GetSubject(mimeContent);
var body = client.GetHtmlBody(mimeContent);
var attachmentCount = client.GetAttachmentCount(mimeContent);
```

### Configuration File Support

Add to `app.config` or `web.config`:

```xml
<configuration>
  <appSettings>
    <add key="CloudMailKit.TenantId" value="your-tenant-id" />
    <add key="CloudMailKit.ClientId" value="your-client-id" />
    <add key="CloudMailKit.ClientSecret" value="your-client-secret" />
    <add key="CloudMailKit.MailboxAddress" value="user@domain.com" />
  </appSettings>
</configuration>
```

Then use:

```csharp
var client = new UnifiedMailClient();
client.InitializeFromConfig();
client.SendSimple("sender@domain.com", "recipient@domain.com",
                  "Subject", "Body", isHtml: false);
```

## Azure AD App Registration

To use CloudMailKit, you need to register an app in Azure AD:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Enter name and supported account types
4. Click "Register"
5. Note the **Application (client) ID** and **Directory (tenant) ID**
6. Go to "Certificates & secrets" → "New client secret"
7. Copy the **secret value** (save it immediately!)
8. Go to "API permissions" → "Add a permission"
9. Choose "Microsoft Graph" → "Application permissions"
10. Add these permissions:
    - `Mail.Send` (to send emails)
    - `Mail.ReadWrite` (to read/manage emails)
    - `User.Read.All` (to access mailboxes)
11. Click "Grant admin consent"

## API Reference

### CloudMailKit.MailKit Namespace

Perfect drop-in replacements for MailKit/MimeKit:

- **SmtpClient** - Send emails via Graph API (replaces MailKit.Net.Smtp.SmtpClient)
- **MimeMessage** - Email message (replaces MimeKit.MimeMessage)
- **MailboxAddress** - Email address (replaces MimeKit.MailboxAddress)
- **InternetAddress** - Base address class (replaces MimeKit.InternetAddress)
- **InternetAddressList** - Address collection (replaces MimeKit.InternetAddressList)
- **BodyBuilder** - Build message bodies (replaces MimeKit.BodyBuilder)
- **TextPart** - Text/HTML content (replaces MimeKit.TextPart)
- **Multipart** - Multipart content (replaces MimeKit.Multipart)
- **MimePart** - Attachments (replaces MimeKit.MimePart)

### CloudMailKit Namespace

Native CloudMailKit API for reading emails:

- **UnifiedMailClient** - Unified client for sending and reading emails
- **GraphFolder** - Email folder from Graph API
- **GraphMessage** - Email message from Graph API
- **MailMessage** - Simple message class for sending

## Compatibility

- **.NET 6.0+** - Primary target
- **COM Visible** - Works with VB6, Magic xpa, Classic ASP
- **100% MailKit API** - All core sending features supported

## What's Supported

✅ Text emails
✅ HTML emails
✅ Mixed text/HTML (alternative)
✅ Attachments
✅ Multiple recipients (To, Cc, Bcc)
✅ Importance/Priority
✅ Custom headers
✅ Email reading via Graph API
✅ Folder management
✅ Message search
✅ MIME parsing

## What's Not Supported

❌ S/MIME encryption
❌ PGP/GPG signing
❌ Advanced MIME features (embedded objects, custom encodings)
❌ IMAP/POP3 protocols
❌ Direct SMTP sending

## License

MIT License - See LICENSE file for details

## Contributing

Contributions welcome! Please open an issue or PR.

## Support

For issues, questions, or suggestions:
- GitHub Issues: https://github.com/PSScript/CloudMailKit/issues
- Email: support@cloudmailkit.com

---

**CloudMailKit** - Making Microsoft Graph email as simple as SMTP 📧
