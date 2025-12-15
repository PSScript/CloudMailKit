using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;

namespace CloudMailKit
{
    /// <summary>
    /// Graph API-based mail reader
    /// </summary>
    [ComVisible(true)]
    [Guid("C8D9E0F1-A2B3-4567-HIJK-789012345EF4")]
    internal class GraphMailReader : IDisposable
    {
        private string _tenantId;
        private string _clientId;
        private string _clientSecret;
        private string _mailboxAddress;
        private HttpClient _httpClient;

        public GraphMailReader()
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

        public string GetMailboxAddress()
        {
            return _mailboxAddress;
        }

        public GraphFolder GetInbox()
        {
            return GetFolder("Inbox");
        }

        public GraphFolder GetFolder(string folderName)
        {
            var folders = ListFolders();
            return folders.FirstOrDefault(f => f.DisplayName.Equals(folderName, StringComparison.OrdinalIgnoreCase));
        }

        public List<GraphFolder> ListFolders()
        {
            return ListFoldersAsync().Result;
        }

        private async Task<List<GraphFolder>> ListFoldersAsync()
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/mailFolders";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.GetAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to list folders: {await response.Content.ReadAsStringAsync()}");
            }

            var json = await response.Content.ReadAsStringAsync();
            var data = JsonSerializer.Deserialize<JsonElement>(json);
            var folders = new List<GraphFolder>();

            if (data.TryGetProperty("value", out var values))
            {
                foreach (var item in values.EnumerateArray())
                {
                    folders.Add(new GraphFolder
                    {
                        Id = GetStringProperty(item, "id"),
                        DisplayName = GetStringProperty(item, "displayName"),
                        ParentFolderId = GetStringProperty(item, "parentFolderId"),
                        ChildFolderCount = GetIntProperty(item, "childFolderCount"),
                        UnreadItemCount = GetIntProperty(item, "unreadItemCount"),
                        TotalItemCount = GetIntProperty(item, "totalItemCount")
                    });
                }
            }

            return folders;
        }

        public List<GraphMessage> ListMessages(string folderId, int maxCount = 100)
        {
            return ListMessagesAsync(folderId, maxCount).Result;
        }

        private async Task<List<GraphMessage>> ListMessagesAsync(string folderId, int maxCount)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/mailFolders/{folderId}/messages?$top={maxCount}";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.GetAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to list messages: {await response.Content.ReadAsStringAsync()}");
            }

            var json = await response.Content.ReadAsStringAsync();
            return ParseMessages(json);
        }

        public string GetMessageMime(string messageId)
        {
            return GetMessageMimeAsync(messageId).Result;
        }

        private async Task<string> GetMessageMimeAsync(string messageId)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages/{messageId}/$value";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.GetAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to get MIME: {await response.Content.ReadAsStringAsync()}");
            }

            return await response.Content.ReadAsStringAsync();
        }

        public GraphMessage GetMessage(string messageId)
        {
            return GetMessageAsync(messageId).Result;
        }

        private async Task<GraphMessage> GetMessageAsync(string messageId)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages/{messageId}";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.GetAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to get message: {await response.Content.ReadAsStringAsync()}");
            }

            var json = await response.Content.ReadAsStringAsync();
            var messages = ParseMessages($"{{\"value\": [{json}]}}");
            return messages.FirstOrDefault();
        }

        public List<GraphMessage> SearchMessages(string query, string folderId = null)
        {
            return SearchMessagesAsync(query, folderId).Result;
        }

        private async Task<List<GraphMessage>> SearchMessagesAsync(string query, string folderId)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);

            string uri;
            if (!string.IsNullOrEmpty(folderId))
            {
                uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/mailFolders/{folderId}/messages?$search=\"{query}\"";
            }
            else
            {
                uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages?$search=\"{query}\"";
            }

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.GetAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to search messages: {await response.Content.ReadAsStringAsync()}");
            }

            var json = await response.Content.ReadAsStringAsync();
            return ParseMessages(json);
        }

        public void MarkAsRead(string messageId)
        {
            UpdateMessageReadStatus(messageId, true);
        }

        public void MarkAsUnread(string messageId)
        {
            UpdateMessageReadStatus(messageId, false);
        }

        private void UpdateMessageReadStatus(string messageId, bool isRead)
        {
            UpdateMessageReadStatusAsync(messageId, isRead).Wait();
        }

        private async Task UpdateMessageReadStatusAsync(string messageId, bool isRead)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages/{messageId}";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var updateBody = new { isRead = isRead };
            var json = JsonSerializer.Serialize(updateBody);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            var response = await _httpClient.PatchAsync(uri, content);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to update read status: {await response.Content.ReadAsStringAsync()}");
            }
        }

        public string MoveMessage(string messageId, string destinationFolderId)
        {
            return MoveMessageAsync(messageId, destinationFolderId).Result;
        }

        private async Task<string> MoveMessageAsync(string messageId, string destinationFolderId)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages/{messageId}/move";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var moveBody = new { destinationId = destinationFolderId };
            var json = JsonSerializer.Serialize(moveBody);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(uri, content);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to move message: {await response.Content.ReadAsStringAsync()}");
            }

            var resultJson = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<JsonElement>(resultJson);
            return GetStringProperty(result, "id");
        }

        public void DeleteMessage(string messageId)
        {
            DeleteMessageAsync(messageId).Wait();
        }

        private async Task DeleteMessageAsync(string messageId)
        {
            var token = await TokenManager.GetAccessTokenAsync(_clientId, _tenantId, _clientSecret);
            var uri = $"https://graph.microsoft.com/v1.0/users/{_mailboxAddress}/messages/{messageId}";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await _httpClient.DeleteAsync(uri);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to delete message: {await response.Content.ReadAsStringAsync()}");
            }
        }

        private List<GraphMessage> ParseMessages(string json)
        {
            var messages = new List<GraphMessage>();
            var data = JsonSerializer.Deserialize<JsonElement>(json);

            if (data.TryGetProperty("value", out var values))
            {
                foreach (var item in values.EnumerateArray())
                {
                    var msg = new GraphMessage
                    {
                        Id = GetStringProperty(item, "id"),
                        Subject = GetStringProperty(item, "subject"),
                        BodyPreview = GetStringProperty(item, "bodyPreview"),
                        IsRead = GetBoolProperty(item, "isRead"),
                        IsDraft = GetBoolProperty(item, "isDraft"),
                        Importance = GetStringProperty(item, "importance"),
                        HasAttachments = GetBoolProperty(item, "hasAttachments"),
                        InternetMessageId = GetStringProperty(item, "internetMessageId"),
                        ConversationId = GetStringProperty(item, "conversationId")
                    };

                    // Parse body
                    if (item.TryGetProperty("body", out var body))
                    {
                        msg.BodyContent = GetStringProperty(body, "content");
                        msg.BodyContentType = GetStringProperty(body, "contentType");
                    }

                    // Parse from
                    if (item.TryGetProperty("from", out var from))
                    {
                        if (from.TryGetProperty("emailAddress", out var fromEmail))
                        {
                            msg.From = GetStringProperty(fromEmail, "address");
                        }
                    }

                    // Parse sender
                    if (item.TryGetProperty("sender", out var sender))
                    {
                        if (sender.TryGetProperty("emailAddress", out var senderEmail))
                        {
                            msg.Sender = GetStringProperty(senderEmail, "address");
                        }
                    }

                    // Parse recipients
                    if (item.TryGetProperty("toRecipients", out var toRecipients))
                    {
                        foreach (var recipient in toRecipients.EnumerateArray())
                        {
                            if (recipient.TryGetProperty("emailAddress", out var email))
                            {
                                msg.ToRecipients.Add(GetStringProperty(email, "address"));
                            }
                        }
                    }

                    if (item.TryGetProperty("ccRecipients", out var ccRecipients))
                    {
                        foreach (var recipient in ccRecipients.EnumerateArray())
                        {
                            if (recipient.TryGetProperty("emailAddress", out var email))
                            {
                                msg.CcRecipients.Add(GetStringProperty(email, "address"));
                            }
                        }
                    }

                    if (item.TryGetProperty("bccRecipients", out var bccRecipients))
                    {
                        foreach (var recipient in bccRecipients.EnumerateArray())
                        {
                            if (recipient.TryGetProperty("emailAddress", out var email))
                            {
                                msg.BccRecipients.Add(GetStringProperty(email, "address"));
                            }
                        }
                    }

                    // Parse dates
                    if (item.TryGetProperty("receivedDateTime", out var receivedDateTime))
                    {
                        if (receivedDateTime.ValueKind == JsonValueKind.String)
                        {
                            DateTime.TryParse(receivedDateTime.GetString(), out var dt);
                            msg.ReceivedDateTime = dt;
                        }
                    }

                    if (item.TryGetProperty("sentDateTime", out var sentDateTime))
                    {
                        if (sentDateTime.ValueKind == JsonValueKind.String)
                        {
                            DateTime.TryParse(sentDateTime.GetString(), out var dt);
                            msg.SentDateTime = dt;
                        }
                    }

                    messages.Add(msg);
                }
            }

            return messages;
        }

        private string GetStringProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var property))
            {
                if (property.ValueKind == JsonValueKind.String)
                {
                    return property.GetString();
                }
            }
            return string.Empty;
        }

        private int GetIntProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var property))
            {
                if (property.ValueKind == JsonValueKind.Number)
                {
                    return property.GetInt32();
                }
            }
            return 0;
        }

        private bool GetBoolProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var property))
            {
                if (property.ValueKind == JsonValueKind.True || property.ValueKind == JsonValueKind.False)
                {
                    return property.GetBoolean();
                }
            }
            return false;
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
