using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CloudMailKit
{
    /// <summary>
    /// Manages OAuth tokens for Microsoft Graph API
    /// Implements token caching and refresh logic
    /// </summary>
    internal static class TokenManager
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private static readonly Dictionary<string, TokenCache> _tokenCache = new Dictionary<string, TokenCache>();
        private static readonly object _lockObject = new object();

        private class TokenCache
        {
            public string AccessToken { get; set; }
            public DateTime ExpiresAt { get; set; }
        }

        public static async Task<string> GetAccessTokenAsync(string clientId, string tenantId, string clientSecret)
        {
            var cacheKey = $"{tenantId}:{clientId}";

            lock (_lockObject)
            {
                if (_tokenCache.TryGetValue(cacheKey, out var cached))
                {
                    // Return cached token if it's still valid (with 5 minute buffer)
                    if (cached.ExpiresAt > DateTime.UtcNow.AddMinutes(5))
                    {
                        return cached.AccessToken;
                    }
                }
            }

            // Request new token
            var token = await RequestTokenAsync(clientId, tenantId, clientSecret);

            lock (_lockObject)
            {
                _tokenCache[cacheKey] = token;
            }

            return token.AccessToken;
        }

        private static async Task<TokenCache> RequestTokenAsync(string clientId, string tenantId, string clientSecret)
        {
            var tokenEndpoint = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            var body = new Dictionary<string, string>
            {
                { "client_id", clientId },
                { "client_secret", clientSecret },
                { "scope", "https://graph.microsoft.com/.default" },
                { "grant_type", "client_credentials" }
            };

            var content = new FormUrlEncodedContent(body);
            var response = await _httpClient.PostAsync(tokenEndpoint, content);

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new Exception($"Token request failed: {response.StatusCode} - {error}");
            }

            var json = await response.Content.ReadAsStringAsync();
            var tokenResponse = JsonSerializer.Deserialize<JsonElement>(json);

            var accessToken = tokenResponse.GetProperty("access_token").GetString();
            var expiresIn = tokenResponse.GetProperty("expires_in").GetInt32();

            return new TokenCache
            {
                AccessToken = accessToken,
                ExpiresAt = DateTime.UtcNow.AddSeconds(expiresIn)
            };
        }

        public static void ClearCache()
        {
            lock (_lockObject)
            {
                _tokenCache.Clear();
            }
        }
    }
}
