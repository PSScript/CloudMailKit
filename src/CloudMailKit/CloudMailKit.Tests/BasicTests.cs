using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace CloudMailKit.Tests
{
    [TestClass]
    public class BasicTests
    {
        private static string _tenantId;
        private static string _clientId;
        private static string _clientSecret;
        private static string _mailbox;

        [ClassInitialize]
        public static void Setup(TestContext context)
        {
            // Load from environment or config
            _tenantId = Environment.GetEnvironmentVariable("TEST_TENANT_ID") 
                ?? "your-tenant-id";
            _clientId = Environment.GetEnvironmentVariable("TEST_CLIENT_ID") 
                ?? "your-client-id";
            _clientSecret = Environment.GetEnvironmentVariable("TEST_CLIENT_SECRET") 
                ?? "your-secret";
            _mailbox = Environment.GetEnvironmentVariable("TEST_MAILBOX") 
                ?? "test@domain.com";
        }

        [TestMethod]
        public void Test_Initialize()
        {
            using (var client = new UnifiedMailClient())
            {
                client.Initialize(_tenantId, _clientId, _clientSecret, _mailbox);
                // Should not throw
            }
        }

        [TestMethod]
        public void Test_GetInbox()
        {
            using (var client = new UnifiedMailClient())
            {
                client.Initialize(_tenantId, _clientId, _clientSecret, _mailbox);
                
                var inbox = client.GetInbox();
                
                Assert.IsNotNull(inbox);
                Assert.AreEqual("Inbox", inbox.DisplayName);
                Assert.IsTrue(inbox.TotalItemCount >= 0);
            }
        }

        [TestMethod]
        public void Test_SendEmail()
        {
            using (var client = new UnifiedMailClient())
            {
                client.Initialize(_tenantId, _clientId, _clientSecret, _mailbox);
                
                client.SendSimple(
                    from: _mailbox,
                    to: _mailbox,
                    subject: "Test Email",
                    body: "This is a test email from CloudMailKit",
                    isHtml: false
                );
                
                // Should not throw
            }
        }
    }
}
```

**Run tests:**
```
Test → Run All Tests (Ctrl+R, A)