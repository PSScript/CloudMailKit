# Legacy Application Rescue Guide

**Mission:** Keep your stone age applications working when Microsoft kills SMTP/basic auth.

## The Problem

Your old applications are about to BREAK:

- ❌ Microsoft 365 is **disabling SMTP with basic authentication**
- ❌ Apps built 10-20 years ago will **stop sending email**
- ❌ Full rewrite = **6-12 months** + $$$
- ❌ Business can't wait that long

## The Solution: CloudMailKit

✅ **Plug in CloudMailKit** - works in minutes, not months
✅ **Keep old code running** - minimal changes required
✅ **Buy time** - modernize at your own pace
✅ **Works with:**
- VB6
- Classic ASP
- .NET Framework 1.1, 2.0, 3.5, 4.0, 4.5, 4.6, 4.7, 4.8
- Magic xpa
- PowerBuilder
- Any COM-capable language

---

## Quick Start by Platform

### VB6 / Classic ASP / COM Applications

**1. Register the DLL:**
```batch
cd C:\path\to\CloudMailKit
regasm CloudMailKit.dll /tlb:CloudMailKit.tlb /codebase
```

**2. VB6 Code:**
```vb
' Old code (BROKEN - SMTP disabled)
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "sender@company.com"
objEmail.To = "recipient@example.com"
objEmail.Subject = "Test Email"
objEmail.TextBody = "Hello from VB6"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "user@company.com"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
objEmail.Send

' New code (WORKS - Graph API)
Set client = CreateObject("CloudMailKit.UnifiedMailClient")
client.Initialize "your-tenant-id", "your-client-id", "your-client-secret", "sender@company.com"
client.SendSimple "sender@company.com", "recipient@example.com", "Test Email", "Hello from VB6", False
```

**3. Classic ASP Code:**
```asp
<%
' Old code (BROKEN)
Set objEmail = Server.CreateObject("CDO.Message")
objEmail.From = "noreply@company.com"
objEmail.To = "customer@example.com"
objEmail.Subject = "Order Confirmation"
objEmail.HTMLBody = "<h1>Thank you for your order!</h1>"
' ... SMTP config ...
objEmail.Send

' New code (WORKS)
Set client = Server.CreateObject("CloudMailKit.UnifiedMailClient")
client.Initialize "tenant-id", "client-id", "secret", "noreply@company.com"
client.SendSimple "noreply@company.com", "customer@example.com", "Order Confirmation", "<h1>Thank you!</h1>", True
Set client = Nothing
%>
```

### Old .NET Framework (1.1 - 4.8)

**1. Install via NuGet:**
```powershell
Install-Package CloudMailKit
```

**2. Old System.Net.Mail code (BROKEN):**
```csharp
using System.Net.Mail;

// This BREAKS when Microsoft disables basic auth
var message = new MailMessage();
message.From = new MailAddress("sender@company.com");
message.To.Add("recipient@example.com");
message.Subject = "Test";
message.Body = "Hello";

var smtp = new SmtpClient("smtp.office365.com", 587);
smtp.EnableSsl = true;
smtp.Credentials = new NetworkCredential("user@company.com", "password");
smtp.Send(message);
```

**3. New CloudMailKit code (WORKS):**
```csharp
using CloudMailKit;

// Quick fix - use simple API
var client = new UnifiedMailClient();
client.Initialize("tenant-id", "client-id", "client-secret", "sender@company.com");
client.SendSimple("sender@company.com", "recipient@example.com", "Test", "Hello", false);

// OR use MailKit-compatible API for drop-in replacement
using CloudMailKit.MailKit;

var message = new MimeMessage();
message.From.Add(new MailboxAddress("Sender", "sender@company.com"));
message.To.Add(new MailboxAddress("Recipient", "recipient@example.com"));
message.Subject = "Test";
message.Body = new TextPart("plain") { Text = "Hello" };

using (var smtpClient = new SmtpClient())
{
    smtpClient.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
    smtpClient.Authenticate("client-id|tenant-id|sender@company.com", "client-secret");
    smtpClient.Send(message);
    smtpClient.Disconnect(true);
}
```

### Magic xpa

**1. Define External Function:**
```
Function Name: SendEmail
DLL Name: CloudMailKit.dll
Class Name: CloudMailKit.UnifiedMailClient
Return Type: Void
Parameters:
  - tenantId (String)
  - clientId (String)
  - clientSecret (String)
  - mailboxAddress (String)
  - from (String)
  - to (String)
  - subject (String)
  - body (String)
  - isHtml (Boolean)
```

**2. Call in Magic:**
```
Call SendEmail(
  'tenant-id',
  'client-id',
  'client-secret',
  'sender@company.com',
  'sender@company.com',
  'recipient@example.com',
  'Test Email',
  'Email body',
  False
)
```

---

## Azure AD Setup (One-Time, 5 Minutes)

### Step 1: Register App
1. Go to https://portal.azure.com
2. Azure Active Directory → App registrations → New registration
3. Name: "Legacy Email Bridge"
4. Register

### Step 2: Get Credentials
- Copy **Application (client) ID** → This is your `client-id`
- Copy **Directory (tenant) ID** → This is your `tenant-id`

### Step 3: Create Secret
1. Certificates & secrets → New client secret
2. Description: "Email sending"
3. Copy the **Value** immediately → This is your `client-secret`

### Step 4: Grant Permissions
1. API permissions → Add a permission
2. Microsoft Graph → Application permissions
3. Select: **Mail.Send**
4. Click "Grant admin consent for [Your Org]"

### Done! ✅
Use these three values in your code:
- `tenant-id`
- `client-id`
- `client-secret`

---

## Installation Options

### Option 1: NuGet (Easiest for .NET)
```powershell
Install-Package CloudMailKit
```

### Option 2: Manual DLL (For VB6/Classic ASP/COM)
1. Download `CloudMailKit.dll`
2. Copy to: `C:\CloudMailKit\`
3. Register as admin:
```batch
cd C:\CloudMailKit
regasm CloudMailKit.dll /tlb /codebase
```

### Option 3: GAC Installation (Enterprise)
```batch
gacutil /i CloudMailKit.dll
regasm CloudMailKit.dll /tlb
```

---

## Configuration File (Recommended)

### For .NET Framework Apps
Add to `web.config` or `app.config`:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="CloudMailKit.TenantId" value="your-tenant-id-here" />
    <add key="CloudMailKit.ClientId" value="your-client-id-here" />
    <add key="CloudMailKit.ClientSecret" value="your-client-secret-here" />
    <add key="CloudMailKit.MailboxAddress" value="noreply@yourcompany.com" />
  </appSettings>
</configuration>
```

Then in code:
```csharp
var client = new UnifiedMailClient();
client.InitializeFromConfig(); // Reads from config file
client.SendSimple(from, to, subject, body, isHtml);
```

### For VB6/Classic ASP
Create `CloudMailKit.ini` in app directory:
```ini
[CloudMailKit]
TenantId=your-tenant-id
ClientId=your-client-id
ClientSecret=your-client-secret
MailboxAddress=noreply@company.com
```

---

## Common Scenarios

### Scenario 1: ASP.NET WebForms (Framework 4.8)

**Before (BROKEN):**
```csharp
protected void btnSend_Click(object sender, EventArgs e)
{
    var mail = new MailMessage("noreply@company.com", txtEmail.Text);
    mail.Subject = "Contact Form Submission";
    mail.Body = txtMessage.Text;

    var smtp = new SmtpClient("smtp.office365.com", 587);
    smtp.Credentials = new NetworkCredential("noreply@company.com", "password");
    smtp.EnableSsl = true;
    smtp.Send(mail);
}
```

**After (WORKS):**
```csharp
using CloudMailKit;

protected void btnSend_Click(object sender, EventArgs e)
{
    var client = new UnifiedMailClient();
    client.InitializeFromConfig(); // Credentials from web.config
    client.SendSimple(
        from: "noreply@company.com",
        to: txtEmail.Text,
        subject: "Contact Form Submission",
        body: txtMessage.Text,
        isHtml: false
    );
}
```

### Scenario 2: Windows Service (.NET Framework 4.5)

```csharp
using CloudMailKit;
using System.Configuration;

public class EmailService : ServiceBase
{
    private UnifiedMailClient _emailClient;

    protected override void OnStart(string[] args)
    {
        _emailClient = new UnifiedMailClient();
        _emailClient.InitializeFromConfig();
    }

    public void SendAlert(string recipient, string message)
    {
        _emailClient.SendSimple(
            from: ConfigurationManager.AppSettings["AlertFromAddress"],
            to: recipient,
            subject: "System Alert",
            body: message,
            isHtml: false
        );
    }
}
```

### Scenario 3: Scheduled Task (VBScript)

**scheduled-email.vbs:**
```vbscript
Set client = CreateObject("CloudMailKit.UnifiedMailClient")
client.Initialize "tenant-id", "client-id", "client-secret", "scheduler@company.com"

' Send daily report
client.SendSimple _
    "scheduler@company.com", _
    "manager@company.com", _
    "Daily Report - " & Date, _
    "The daily report is attached.", _
    False

Set client = Nothing
WScript.Echo "Email sent successfully!"
```

---

## Troubleshooting

### Error: "Class not registered"
**Solution:** Run as Administrator:
```batch
regasm CloudMailKit.dll /tlb /codebase
```

### Error: "Could not load file or assembly"
**Solution:** Ensure .NET Framework 4.8 is installed:
```batch
# Download from Microsoft
# https://dotnet.microsoft.com/download/dotnet-framework/net48
```

### Error: "Authentication failed"
**Solution:** Verify Azure AD app:
1. Permissions granted? (Mail.Send)
2. Admin consent granted?
3. Correct tenant/client ID?
4. Secret not expired?

### Error: "Access Denied"
**Solution:** Sender must have mailbox:
- The `from` address must be a real mailbox in your Microsoft 365 tenant
- Can't send from external addresses

---

## Performance & Limits

- **Token Caching:** Automatic - tokens reused for 1 hour
- **Rate Limits:** Microsoft Graph limits apply (~30 msgs/min per mailbox)
- **Attachments:** Up to 4 MB per message
- **Recipients:** Up to 500 per message

---

## Security Best Practices

### ✅ DO:
- Store credentials in config files (encrypted if possible)
- Use app.config encryption:
  ```batch
  aspnet_regiis -pe "appSettings" -app "/YourApp"
  ```
- Use dedicated service account mailbox
- Restrict Azure AD app to Mail.Send only
- Monitor usage in Azure AD logs

### ❌ DON'T:
- Hard-code secrets in source code
- Commit secrets to source control
- Use personal mailboxes for apps
- Share client secrets between apps

---

## Migration Strategy

### Phase 1: Emergency Fix (1 day)
1. Register Azure AD app
2. Install CloudMailKit
3. Update 1-2 critical apps
4. Test thoroughly
5. Monitor for issues

### Phase 2: Rollout (1 week)
1. Update remaining apps
2. Centralize configuration
3. Document changes
4. Train support staff

### Phase 3: Modernize (6-12 months)
1. Plan full rewrite with modern stack
2. Consider using AI code assistance
3. Gradually replace old code
4. CloudMailKit keeps things working meanwhile

---

## Support for Stone Age Apps

CloudMailKit is specifically designed to rescue legacy applications:

- ✅ **VB6** - Full COM support
- ✅ **Classic ASP** - Server.CreateObject
- ✅ **ASP.NET WebForms** - .NET Framework 2.0+
- ✅ **Windows Services** - All Framework versions
- ✅ **.NET Framework 1.1-4.8** - Multi-targeted
- ✅ **Magic xpa** - External functions
- ✅ **PowerBuilder** - COM interop
- ✅ **VBScript** - CreateObject
- ✅ **JScript** - ActiveXObject

---

## Real-World Example: E-Commerce Site

**Scenario:** ASP.NET WebForms site from 2008, sends order confirmations via SMTP

**Problem:** Microsoft disables basic auth, emails stop working, orders pile up

**Solution:** 30-minute fix with CloudMailKit

**Before (OrderConfirmation.aspx.cs):**
```csharp
void SendOrderEmail(Order order)
{
    var msg = new MailMessage();
    msg.From = new MailAddress("orders@shop.com");
    msg.To.Add(order.CustomerEmail);
    msg.Subject = "Order #" + order.OrderId;
    msg.Body = GenerateOrderHtml(order);
    msg.IsBodyHtml = true;

    var smtp = new SmtpClient("smtp.office365.com", 587);
    smtp.Credentials = new NetworkCredential("orders@shop.com", "password123");
    smtp.EnableSsl = true;
    smtp.Send(msg); // BREAKS!
}
```

**After:**
```csharp
using CloudMailKit;

// Add to web.config once
void SendOrderEmail(Order order)
{
    var client = new UnifiedMailClient();
    client.InitializeFromConfig(); // Reads from web.config
    client.SendSimple(
        from: "orders@shop.com",
        to: order.CustomerEmail,
        subject: "Order #" + order.OrderId,
        body: GenerateOrderHtml(order),
        isHtml: true
    );
}
```

**Result:** Site back online in 30 minutes. Full rewrite can wait.

---

## Get Help

- 📧 Email: support@cloudmailkit.com
- 💬 GitHub: https://github.com/PSScript/CloudMailKit/issues
- 📚 Docs: https://github.com/PSScript/CloudMailKit

**Remember:** CloudMailKit is your **lifeline** for legacy apps. It buys you time to modernize properly without breaking production! 🚀
