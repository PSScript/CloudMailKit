using Microsoft.VisualStudio.TestTools.UnitTesting;
using CloudMailKit.MailKit;
using System;

namespace CloudMailKit.Tests
{
    /// <summary>
    /// Tests demonstrating 100% MailKit/MimeKit API compatibility
    /// These tests show that CloudMailKit can be used as a drop-in replacement
    /// </summary>
    [TestClass]
    public class MailKitCompatibilityTests
    {
        private static string _tenantId;
        private static string _clientId;
        private static string _clientSecret;
        private static string _senderAddress;

        [ClassInitialize]
        public static void Setup(TestContext context)
        {
            // Load from environment or config
            _tenantId = Environment.GetEnvironmentVariable("TEST_TENANT_ID") ?? "your-tenant-id";
            _clientId = Environment.GetEnvironmentVariable("TEST_CLIENT_ID") ?? "your-client-id";
            _clientSecret = Environment.GetEnvironmentVariable("TEST_CLIENT_SECRET") ?? "your-secret";
            _senderAddress = Environment.GetEnvironmentVariable("TEST_SENDER") ?? "sender@domain.com";
        }

        [TestMethod]
        public void Test_MimeMessage_Creation()
        {
            // This is IDENTICAL to how you would use MimeKit.MimeMessage
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Test Sender", "sender@test.com"));
            message.To.Add(new MailboxAddress("Test Recipient", "recipient@test.com"));
            message.Subject = "Test Subject";

            Assert.IsNotNull(message);
            Assert.AreEqual(1, message.From.Count);
            Assert.AreEqual(1, message.To.Count);
            Assert.AreEqual("Test Subject", message.Subject);
            Assert.AreEqual("sender@test.com", message.From[0].Address);
            Assert.AreEqual("recipient@test.com", message.To[0].Address);
        }

        [TestMethod]
        public void Test_BodyBuilder_TextAndHtml()
        {
            // This is IDENTICAL to how you would use MimeKit.BodyBuilder
            var builder = new BodyBuilder();
            builder.TextBody = "This is plain text";
            builder.HtmlBody = "<p>This is HTML</p>";

            var body = builder.ToMessageBody();

            Assert.IsNotNull(body);
            Assert.IsInstanceOfType(body, typeof(Multipart));

            var multipart = (Multipart)body;
            Assert.IsTrue(multipart.Count >= 2);
        }

        [TestMethod]
        public void Test_BodyBuilder_WithAttachments()
        {
            var builder = new BodyBuilder();
            builder.TextBody = "Email with attachments";

            // Create a temp file for testing
            var tempFile = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllText(tempFile, "Test attachment content");

            try
            {
                builder.AddAttachment(tempFile);

                var body = builder.ToMessageBody();
                Assert.IsNotNull(body);
                Assert.IsInstanceOfType(body, typeof(Multipart));

                var multipart = (Multipart)body;
                Assert.IsTrue(multipart.Count >= 2); // Text + attachment
            }
            finally
            {
                System.IO.File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void Test_MailboxAddress_Parsing()
        {
            // This is IDENTICAL to how you would use MimeKit.MailboxAddress
            var addr1 = new MailboxAddress("John Doe", "john@example.com");
            Assert.AreEqual("John Doe", addr1.Name);
            Assert.AreEqual("john@example.com", addr1.Address);

            var addr2 = MailboxAddress.Parse("Jane Smith <jane@example.com>");
            Assert.AreEqual("Jane Smith", addr2.Name);
            Assert.AreEqual("jane@example.com", addr2.Address);

            var addr3 = MailboxAddress.Parse("simple@example.com");
            Assert.AreEqual("simple@example.com", addr3.Address);
        }

        [TestMethod]
        public void Test_InternetAddressList()
        {
            // This is IDENTICAL to how you would use MimeKit.InternetAddressList
            var list = new InternetAddressList();
            list.Add(new MailboxAddress("User 1", "user1@example.com"));
            list.Add(new MailboxAddress("User 2", "user2@example.com"));
            list.Add("user3@example.com");

            Assert.AreEqual(3, list.Count);
            Assert.AreEqual("user1@example.com", list[0].Address);
            Assert.AreEqual("user2@example.com", list[1].Address);
            Assert.AreEqual("user3@example.com", list[2].Address);
        }

        [TestMethod]
        public void Test_TextPart_Plain()
        {
            // This is IDENTICAL to how you would use MimeKit.TextPart
            var part = new TextPart("plain")
            {
                Text = "This is plain text"
            };

            Assert.IsTrue(part.IsPlain);
            Assert.IsFalse(part.IsHtml);
            Assert.AreEqual("This is plain text", part.Text);
        }

        [TestMethod]
        public void Test_TextPart_Html()
        {
            var part = new TextPart("html")
            {
                Text = "<p>This is HTML</p>"
            };

            Assert.IsFalse(part.IsPlain);
            Assert.IsTrue(part.IsHtml);
            Assert.AreEqual("<p>This is HTML</p>", part.Text);
        }

        [TestMethod]
        public void Test_Multipart_Mixed()
        {
            // This is IDENTICAL to how you would use MimeKit.Multipart
            var multipart = new Multipart("mixed");
            multipart.Add(new TextPart("plain") { Text = "Part 1" });
            multipart.Add(new TextPart("html") { Text = "<p>Part 2</p>" });

            Assert.AreEqual(2, multipart.Count);
            Assert.IsInstanceOfType(multipart[0], typeof(TextPart));
            Assert.IsInstanceOfType(multipart[1], typeof(TextPart));
        }

        [TestMethod]
        public void Test_MimeMessage_Importance()
        {
            var message = new MimeMessage();
            message.Subject = "Important Message";
            message.Importance = MessageImportance.High;

            Assert.AreEqual(MessageImportance.High, message.Importance);
        }

        [TestMethod]
        public void Test_MimeMessage_MultipleRecipients()
        {
            var message = new MimeMessage();
            message.To.Add(new MailboxAddress("User 1", "user1@example.com"));
            message.To.Add(new MailboxAddress("User 2", "user2@example.com"));
            message.Cc.Add(new MailboxAddress("CC User", "cc@example.com"));
            message.Bcc.Add(new MailboxAddress("BCC User", "bcc@example.com"));

            Assert.AreEqual(2, message.To.Count);
            Assert.AreEqual(1, message.Cc.Count);
            Assert.AreEqual(1, message.Bcc.Count);
        }

        [TestMethod]
        public void Test_SmtpClient_Connect()
        {
            // This is IDENTICAL to how you would use MailKit.Net.Smtp.SmtpClient
            using (var client = new SmtpClient())
            {
                Assert.IsFalse(client.IsConnected);
                Assert.IsFalse(client.IsAuthenticated);

                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);

                Assert.IsTrue(client.IsConnected);
                Assert.IsFalse(client.IsAuthenticated);

                client.Disconnect(true);
                Assert.IsFalse(client.IsConnected);
            }
        }

        [TestMethod]
        public void Test_SmtpClient_Authenticate()
        {
            using (var client = new SmtpClient())
            {
                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);

                // Authenticate with Graph credentials
                client.Authenticate($"{_clientId}|{_tenantId}|{_senderAddress}", _clientSecret);

                Assert.IsTrue(client.IsAuthenticated);
                client.Disconnect(true);
            }
        }

        [TestMethod]
        public void Test_SmtpClient_SetGraphCredentials()
        {
            using (var client = new SmtpClient())
            {
                client.SetGraphCredentials(_tenantId, _clientId, _clientSecret, _senderAddress);

                Assert.IsTrue(client.IsConnected);
                Assert.IsTrue(client.IsAuthenticated);
            }
        }

        [TestMethod]
        [Ignore("Requires valid Graph API credentials")]
        public void Test_SendEmail_FullExample()
        {
            // This is a COMPLETE example showing 100% MailKit compatibility
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Test Sender", _senderAddress));
            message.To.Add(new MailboxAddress("Test Recipient", _senderAddress)); // Send to self
            message.Subject = "CloudMailKit Test - MailKit Compatibility";

            var builder = new BodyBuilder();
            builder.TextBody = "This is the plain text version of the email.";
            builder.HtmlBody = @"
                <h1>CloudMailKit Test</h1>
                <p>This email was sent using CloudMailKit with <strong>100% MailKit API compatibility</strong>.</p>
                <p>The code is identical to MailKit/MimeKit!</p>
            ";
            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
                client.Authenticate($"{_clientId}|{_tenantId}|{_senderAddress}", _clientSecret);
                client.Send(message);
                client.Disconnect(true);
            }

            // If we get here without exception, the test passed
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Test_MimeMessage_CreateReply()
        {
            var original = new MimeMessage();
            original.From.Add(new MailboxAddress("Alice", "alice@example.com"));
            original.To.Add(new MailboxAddress("Bob", "bob@example.com"));
            original.Subject = "Original Message";
            original.MessageId = "<123@example.com>";
            original.Body = new TextPart("plain") { Text = "Original text" };

            var reply = original.CreateReply(false);

            Assert.AreEqual("Re: Original Message", reply.Subject);
            Assert.AreEqual(1, reply.To.Count);
            Assert.AreEqual("alice@example.com", reply.To[0].Address);
            Assert.AreEqual("<123@example.com>", reply.InReplyTo);
        }
    }
}
