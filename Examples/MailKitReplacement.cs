using CloudMailKit.MailKit;
using System;

namespace CloudMailKit.Examples
{
    /// <summary>
    /// Example showing how to use CloudMailKit as a 100% drop-in replacement for MailKit
    ///
    /// This is IDENTICAL to how you would use MailKit.Net.Smtp.SmtpClient and MimeKit.MimeMessage
    /// The ONLY difference is the namespace: CloudMailKit.MailKit instead of MailKit/MimeKit
    /// </summary>
    public class MailKitReplacementExample
    {
        public static void SendSimpleEmail()
        {
            // Create a message - IDENTICAL to MimeKit.MimeMessage
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Sender Name", "sender@company.com"));
            message.To.Add(new MailboxAddress("Recipient Name", "recipient@example.com"));
            message.Subject = "Test Email from CloudMailKit";

            // Create body - IDENTICAL to MimeKit.BodyBuilder
            var builder = new BodyBuilder();
            builder.TextBody = "This is the plain text version of the email.";
            builder.HtmlBody = "<p>This is the <strong>HTML</strong> version of the email.</p>";
            message.Body = builder.ToMessageBody();

            // Send via Graph API - IDENTICAL API to MailKit.Net.Smtp.SmtpClient
            using (var client = new SmtpClient())
            {
                // Connect (Graph API endpoint - but same method signature as MailKit)
                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);

                // Authenticate with Graph OAuth
                // Format: "client_id|tenant_id|sender@domain.com"
                client.Authenticate(
                    userName: "your-app-id|your-tenant-id|sender@company.com",
                    password: "your-client-secret"
                );

                // Send message - IDENTICAL to MailKit
                client.Send(message);

                // Disconnect - IDENTICAL to MailKit
                client.Disconnect(true);
            }

            Console.WriteLine("Email sent successfully!");
        }

        public static void SendEmailWithAttachments()
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Sales Team", "sales@company.com"));
            message.To.Add(new MailboxAddress("John Doe", "john@example.com"));
            message.Cc.Add(new MailboxAddress("Manager", "manager@company.com"));
            message.Subject = "Monthly Sales Report";
            message.Importance = MessageImportance.High;

            var builder = new BodyBuilder();
            builder.HtmlBody = @"
                <h1>Monthly Sales Report</h1>
                <p>Dear John,</p>
                <p>Please find attached the sales report for this month.</p>
                <p>Best regards,<br/>Sales Team</p>
            ";

            // Add attachments - IDENTICAL to MailKit
            builder.AddAttachment("reports/sales-report.pdf");
            builder.AddAttachment("reports/charts.xlsx");

            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
                client.Authenticate("app-id|tenant-id|sales@company.com", "secret");
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine("Email with attachments sent successfully!");
        }

        public static async System.Threading.Tasks.Task SendEmailAsync()
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Newsletter", "news@company.com"));
            message.To.Add(new MailboxAddress("Subscriber", "subscriber@example.com"));
            message.Subject = "Weekly Newsletter";

            var builder = new BodyBuilder();
            builder.TextBody = "This is the plain text version of our newsletter.";
            builder.HtmlBody = @"
                <h1>Weekly Newsletter</h1>
                <p>Welcome to our weekly newsletter!</p>
                <p>This week's highlights:</p>
                <ul>
                    <li>Product launch</li>
                    <li>Special offers</li>
                    <li>Community updates</li>
                </ul>
            ";
            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                // Alternative authentication method
                client.SetGraphCredentials(
                    tenantId: "your-tenant-id",
                    clientId: "your-client-id",
                    clientSecret: "your-client-secret",
                    senderAddress: "news@company.com"
                );

                // Async send - IDENTICAL to MailKit
                await client.SendAsync(message);
            }

            Console.WriteLine("Newsletter sent asynchronously!");
        }

        public static void SendToMultipleRecipients()
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("HR Department", "hr@company.com"));

            // Add multiple recipients
            message.To.Add(new MailboxAddress("Employee 1", "emp1@company.com"));
            message.To.Add(new MailboxAddress("Employee 2", "emp2@company.com"));
            message.To.Add(new MailboxAddress("Employee 3", "emp3@company.com"));

            // Add CC recipients
            message.Cc.Add(new MailboxAddress("Manager", "manager@company.com"));

            // Add BCC recipients
            message.Bcc.Add(new MailboxAddress("Director", "director@company.com"));

            message.Subject = "Company Policy Update";

            var builder = new BodyBuilder();
            builder.HtmlBody = "<p>Dear Team,</p><p>Please review the updated company policies.</p>";
            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
                client.Authenticate("app-id|tenant-id|hr@company.com", "secret");
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine("Email sent to multiple recipients!");
        }

        public static void MigrationFromMailKit()
        {
            // BEFORE: Using MailKit/MimeKit
            // using MailKit.Net.Smtp;
            // using MimeKit;

            // AFTER: Using CloudMailKit (drop-in replacement)
            // using CloudMailKit.MailKit;

            // ALL THE REST OF THE CODE STAYS EXACTLY THE SAME!

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Sender", "sender@domain.com"));
            message.To.Add(new MailboxAddress("Recipient", "recipient@domain.com"));
            message.Subject = "Migration Test";

            message.Body = new TextPart("plain")
            {
                Text = "This code works with both MailKit and CloudMailKit!"
            };

            using (var client = new SmtpClient())
            {
                // Only difference: authentication credentials format for Graph OAuth
                // Old: client.Authenticate("smtp-username", "smtp-password");
                // New: client.Authenticate("client-id|tenant-id|sender@domain.com", "client-secret");

                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
                client.Authenticate("app-id|tenant-id|sender@domain.com", "secret");
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine("Migration successful!");
        }
    }
}
