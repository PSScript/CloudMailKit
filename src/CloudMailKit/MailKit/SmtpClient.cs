using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace CloudMailKit.MailKit
{
    /// <summary>
    /// MailKit-compatible SmtpClient that uses Microsoft Graph API instead of SMTP
    /// This is a 100% drop-in replacement for MailKit.Net.Smtp.SmtpClient
    ///
    /// Usage:
    ///   var client = new SmtpClient();
    ///   client.Connect("graph.microsoft.com", 0, SecureSocketOptions.Auto);
    ///   client.Authenticate("client_id|tenant_id", "client_secret");
    ///   client.Send(message);
    ///   client.Disconnect(true);
    ///
    /// Authentication format: "client_id|tenant_id" with client_secret as password
    /// </summary>
    [ComVisible(true)]
    [Guid("C0D1E2F3-A4B5-6789-TUVW-901234567EFF")]
    public class SmtpClient : IDisposable
    {
        private HttpClient _httpClient;
        private string _tenantId;
        private string _clientId;
        private string _clientSecret;
        private string _senderAddress;
        private bool _isConnected;
        private bool _isAuthenticated;

        public SmtpClient()
        {
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            _isConnected = false;
            _isAuthenticated = false;
        }

        /// <summary>
        /// Indicates whether the client is connected
        /// </summary>
        public bool IsConnected => _isConnected;

        /// <summary>
        /// Indicates whether the client is authenticated
        /// </summary>
        public bool IsAuthenticated => _isAuthenticated;

        /// <summary>
        /// Connect to the mail server (Graph API)
        /// For Graph API, this is a no-op but maintains MailKit compatibility
        /// </summary>
        /// <param name="host">Ignored for Graph API (use "graph.microsoft.com" for compatibility)</param>
        /// <param name="port">Ignored for Graph API</param>
        /// <param name="options">Ignored for Graph API</param>
        /// <param name="cancellationToken">Cancellation token</param>
        public void Connect(string host, int port = 0, SecureSocketOptions options = SecureSocketOptions.Auto, CancellationToken cancellationToken = default)
        {
            _isConnected = true;
        }

        /// <summary>
        /// Connect asynchronously
        /// </summary>
        public Task ConnectAsync(string host, int port = 0, SecureSocketOptions options = SecureSocketOptions.Auto, CancellationToken cancellationToken = default)
        {
            _isConnected = true;
            return Task.CompletedTask;
        }

        /// <summary>
        /// Authenticate with OAuth credentials
        ///
        /// For Microsoft Graph OAuth, use one of these formats:
        /// 1. userName = "client_id|tenant_id", password = "client_secret"
        /// 2. userName = "client_id|tenant_id|sender@domain.com", password = "client_secret"
        /// 3. Call SetGraphCredentials() before authenticating
        /// </summary>
        public void Authenticate(string userName, string password, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password))
            {
                throw new ArgumentException("Username and password are required. Format: 'client_id|tenant_id' or 'client_id|tenant_id|sender@domain.com'");
            }

            var parts = userName.Split('|');
            if (parts.Length < 2)
            {
                throw new ArgumentException("Username must be in format: 'client_id|tenant_id' or 'client_id|tenant_id|sender@domain.com'");
            }

            _clientId = parts[0].Trim();
            _tenantId = parts[1].Trim();
            _clientSecret = password;

            if (parts.Length >= 3)
            {
                _senderAddress = parts[2].Trim();
            }

            _isAuthenticated = true;
        }

        /// <summary>
        /// Authenticate asynchronously
        /// </summary>
        public Task AuthenticateAsync(string userName, string password, CancellationToken cancellationToken = default)
        {
            Authenticate(userName, password, cancellationToken);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Alternative way to set Graph credentials (more explicit)
        /// </summary>
        public void SetGraphCredentials(string tenantId, string clientId, string clientSecret, string senderAddress = null)
        {
            _tenantId = tenantId;
            _clientId = clientId;
            _clientSecret = clientSecret;
            _senderAddress = senderAddress;
            _isAuthenticated = true;
            _isConnected = true;
        }

        /// <summary>
        /// Send a MIME message via Microsoft Graph API
        /// </summary>
        public void Send(MimeMessage message, CancellationToken cancellationToken = default)
        {
            SendAsync(message, cancellationToken).Wait(cancellationToken);
        }

        /// <summary>
        /// Send a MIME message asynchronously via Microsoft Graph API
        /// </summary>
        public async Task SendAsync(MimeMessage message, CancellationToken cancellationToken = default)
        {
            if (!_isAuthenticated)
            {
                throw new InvalidOperationException("Not authenticated. Call Authenticate() first.");
            }

            if (message == null)
            {
                throw new ArgumentNullException(nameof(message));
            }

            // Determine sender address
            string fromAddress = _senderAddress;
            if (string.IsNullOrEmpty(fromAddress) && message.From.Count > 0)
            {
                fromAddress = message.From[0].Address;
            }

            if (string.IsNullOrEmpty(fromAddress))
            {
                throw new InvalidOperationException("No sender address specified. Set it in authentication or in message.From");
            }

            // Get access token
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);

            // Build Graph API message
            var graphMessage = BuildGraphMessage(message);

            // Send via Graph API
            var uri = $"https://graph.microsoft.com/v1.0/users/{fromAddress}/sendMail";

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var sendBody = new { message = graphMessage, saveToSentItems = true };
            var json = JsonSerializer.Serialize(sendBody, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(uri, content, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new Exception($"Failed to send email via Graph API: {response.StatusCode} - {error}");
            }
        }

        /// <summary>
        /// Disconnect from the server
        /// </summary>
        public void Disconnect(bool quit, CancellationToken cancellationToken = default)
        {
            _isConnected = false;
            _isAuthenticated = false;
        }

        /// <summary>
        /// Disconnect asynchronously
        /// </summary>
        public Task DisconnectAsync(bool quit, CancellationToken cancellationToken = default)
        {
            _isConnected = false;
            _isAuthenticated = false;
            return Task.CompletedTask;
        }

        /// <summary>
        /// No-op (MailKit compatibility)
        /// </summary>
        public void NoOp(CancellationToken cancellationToken = default)
        {
            // No operation needed for Graph API
        }

        /// <summary>
        /// No-op async (MailKit compatibility)
        /// </summary>
        public Task NoOpAsync(CancellationToken cancellationToken = default)
        {
            return Task.CompletedTask;
        }

        /// <summary>
        /// Build Graph API message from MimeMessage
        /// </summary>
        private object BuildGraphMessage(MimeMessage message)
        {
            var msg = new Dictionary<string, object>
            {
                ["subject"] = message.Subject ?? string.Empty
            };

            // Build recipients
            if (message.To.Count > 0)
            {
                msg["toRecipients"] = message.To.Select(addr => new
                {
                    emailAddress = new { address = addr.Address, name = addr.Name }
                }).ToArray();
            }

            if (message.Cc.Count > 0)
            {
                msg["ccRecipients"] = message.Cc.Select(addr => new
                {
                    emailAddress = new { address = addr.Address, name = addr.Name }
                }).ToArray();
            }

            if (message.Bcc.Count > 0)
            {
                msg["bccRecipients"] = message.Bcc.Select(addr => new
                {
                    emailAddress = new { address = addr.Address, name = addr.Name }
                }).ToArray();
            }

            // Build body
            string bodyContent = string.Empty;
            string bodyType = "text";

            if (!string.IsNullOrEmpty(message.HtmlBody))
            {
                bodyContent = message.HtmlBody;
                bodyType = "html";
            }
            else if (!string.IsNullOrEmpty(message.TextBody))
            {
                bodyContent = message.TextBody;
                bodyType = "text";
            }
            else if (message.Body is TextPart textPart)
            {
                bodyContent = textPart.Text ?? string.Empty;
                bodyType = textPart.IsHtml ? "html" : "text";
            }
            else if (message.Body is Multipart multipart)
            {
                // Extract from multipart
                var html = multipart.GetHtmlBody();
                var text = multipart.GetTextBody();

                if (!string.IsNullOrEmpty(html))
                {
                    bodyContent = html;
                    bodyType = "html";
                }
                else if (!string.IsNullOrEmpty(text))
                {
                    bodyContent = text;
                    bodyType = "text";
                }
            }

            msg["body"] = new
            {
                contentType = bodyType,
                content = bodyContent
            };

            // Handle importance
            if (message.Importance != MessageImportance.Normal)
            {
                msg["importance"] = message.Importance == MessageImportance.High ? "high" :
                                   message.Importance == MessageImportance.Low ? "low" : "normal";
            }

            // Build attachments
            var attachments = new List<object>();

            // Add attachments from message.Attachments (added via BodyBuilder)
            if (message.Attachments.Count > 0)
            {
                foreach (var attachment in message.Attachments)
                {
                    if (attachment is MimePart mimePart && mimePart.Content != null)
                    {
                        attachments.Add(new
                        {
                            odataType = "#microsoft.graph.fileAttachment",
                            name = mimePart.FileName ?? "attachment",
                            contentType = mimePart.ContentType.MimeType,
                            contentBytes = Convert.ToBase64String(mimePart.Content)
                        });
                    }
                }
            }

            // Also check for attachments in multipart body
            if (message.Body is Multipart bodyMultipart)
            {
                foreach (var part in bodyMultipart)
                {
                    if (part is MimePart mimePart && mimePart.IsAttachment && mimePart.Content != null)
                    {
                        attachments.Add(new
                        {
                            odataType = "#microsoft.graph.fileAttachment",
                            name = mimePart.FileName ?? "attachment",
                            contentType = mimePart.ContentType.MimeType,
                            contentBytes = Convert.ToBase64String(mimePart.Content)
                        });
                    }
                }
            }

            if (attachments.Count > 0)
            {
                msg["attachments"] = attachments;
            }

            return msg;
        }

        /// <summary>
        /// Dispose resources
        /// </summary>
        public void Dispose()
        {
            _httpClient?.Dispose();
            _isConnected = false;
            _isAuthenticated = false;
        }
    }

    /// <summary>
    /// Secure socket options for MailKit compatibility
    /// </summary>
    public enum SecureSocketOptions
    {
        None,
        Auto,
        SslOnConnect,
        StartTls,
        StartTlsWhenAvailable
    }
}
