using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CloudMailKit
{
    /// <summary>
    /// Graph API-based mail sender
    /// </summary>
    [ComVisible(true)]
    [Guid("B8C9D0E1-F2A3-4567-HIJK-678901234EF1")]
    internal class GraphMailSender : IDisposable
    {
        private string _tenantId;
        private string _clientId;
        private string _clientSecret;
        private string _mailboxAddress;
        private HttpClient _httpClient;

        public GraphMailSender()
        {
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        }

        public void Initialize(string tenantId, string clientId, string clientSecret, string mailboxAddress)
        {
            _tenantId = tenantId;
            _clientId = clientId;
            _clientSecret = clientSecret;
            _mailboxAddress = mailboxAddress;
        }

        public void SendMessage(MailMessage message)
        {
            SendMessageAsync(message).Wait();
        }

        public void SendSimple(string from, string to, string subject, string body, bool isHtml)
        {
            var msg = new MailMessage
            {
                From = from,
                Subject = subject,
                Body = body,
                IsHtml = isHtml
            };
            msg.AddTo(to);

            SendMessage(msg);
        }

        public void SendWithAttachment(string from, string to, string subject, string body, string attachmentPath)
        {
            var msg = new MailMessage
            {
                From = from,
                Subject = subject,
                Body = body,
                IsHtml = false
            };
            msg.AddTo(to);
            msg.AddAttachment(attachmentPath);

            SendMessage(msg);
        }

        private async Task SendMessageAsync(MailMessage message)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);

            var graphMessage = BuildGraphMessage(message);

            var uri = $"https://graph.microsoft.com/v1.0/users/{message.From}/sendMail";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var sendBody = new { message = graphMessage, saveToSentItems = true };
            var json = JsonSerializer.Serialize(sendBody, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(uri, content);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Send failed: {await response.Content.ReadAsStringAsync()}");
            }
        }

        private object BuildGraphMessage(MailMessage message)
        {
            var msg = new Dictionary<string, object>
            {
                ["subject"] = message.Subject ?? "",
                ["body"] = new
                {
                    contentType = message.IsHtml ? "HTML" : "Text",
                    content = message.Body ?? ""
                },
                ["toRecipients"] = message.To.Select(email => new
                {
                    emailAddress = new { address = email }
                }).ToArray()
            };

            if (message.Cc.Count > 0)
            {
                msg["ccRecipients"] = message.Cc.Select(email => new
                {
                    emailAddress = new { address = email }
                }).ToArray();
            }

            if (message.Bcc.Count > 0)
            {
                msg["bccRecipients"] = message.Bcc.Select(email => new
                {
                    emailAddress = new { address = email }
                }).ToArray();
            }

            if (!string.IsNullOrEmpty(message.Importance))
            {
                msg["importance"] = message.Importance.ToLower();
            }

            if (message.Attachments.Count > 0)
            {
                var attachments = new List<object>();
                foreach (var path in message.Attachments)
                {
                    if (File.Exists(path))
                    {
                        var bytes = File.ReadAllBytes(path);
                        var base64 = Convert.ToBase64String(bytes);
                        var fileName = Path.GetFileName(path);

                        attachments.Add(new
                        {
                            odataType = "#microsoft.graph.fileAttachment",
                            name = fileName,
                            contentBytes = base64
                        });
                    }
                }

                if (attachments.Count > 0)
                {
                    msg["attachments"] = attachments;
                }
            }

            return msg;
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}