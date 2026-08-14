using System.Net.Http.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.HubSpot;

/// <summary>
/// Supplies a HubSpot bearer token. Preferred path: exchange the configured OAuth refresh token for a
/// short-lived access token (cached, refreshed before expiry). Falls back to a static private-app token
/// if no refresh token is configured.
/// </summary>
public sealed class HubSpotTokenProvider
{
    private readonly HttpClient _http;
    private readonly HubSpotOptions _options;
    private readonly ILogger<HubSpotTokenProvider> _logger;
    private readonly SemaphoreSlim _lock = new(1, 1);

    private string? _cachedToken;
    private DateTimeOffset _expiresAtUtc = DateTimeOffset.MinValue;

    public HubSpotTokenProvider(HttpClient http, IOptions<HubSpotOptions> options, ILogger<HubSpotTokenProvider> logger)
    {
        _http = http;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
    {
        // Fallback: a static private-app token when OAuth isn't configured.
        if (string.IsNullOrWhiteSpace(_options.RefreshToken))
        {
            if (!string.IsNullOrWhiteSpace(_options.AccessToken))
            {
                return _options.AccessToken;
            }

            throw new InvalidOperationException(
                "No HubSpot credentials configured. Set HubSpot:RefreshToken (+ClientId/ClientSecret) or HubSpot:AccessToken.");
        }

        if (_cachedToken is not null && DateTimeOffset.UtcNow < _expiresAtUtc.AddSeconds(-60))
        {
            return _cachedToken;
        }

        await _lock.WaitAsync(cancellationToken);
        try
        {
            if (_cachedToken is not null && DateTimeOffset.UtcNow < _expiresAtUtc.AddSeconds(-60))
            {
                return _cachedToken;
            }

            var tokenUrl = $"{_options.BaseUrl.TrimEnd('/')}/oauth/v1/token";
            var form = new Dictionary<string, string>
            {
                ["grant_type"] = "refresh_token",
                ["client_id"] = _options.ClientId,
                ["client_secret"] = _options.ClientSecret,
                ["refresh_token"] = _options.RefreshToken,
            };

            using var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl)
            {
                Content = new FormUrlEncodedContent(form),
            };
            using var response = await _http.SendAsync(request, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                throw new InvalidOperationException(
                    $"HubSpot token refresh failed ({(int)response.StatusCode}): {body}");
            }

            var token = await response.Content.ReadFromJsonAsync<TokenResponse>(cancellationToken)
                ?? throw new InvalidOperationException("HubSpot token response was empty.");

            _cachedToken = token.AccessToken;
            _expiresAtUtc = DateTimeOffset.UtcNow.AddSeconds(token.ExpiresIn);
            _logger.LogDebug("Acquired HubSpot access token, expires in {Seconds}s.", token.ExpiresIn);
            return _cachedToken!;
        }
        finally
        {
            _lock.Release();
        }
    }

    private sealed record TokenResponse
    {
        [JsonPropertyName("access_token")]
        public string AccessToken { get; init; } = string.Empty;

        [JsonPropertyName("expires_in")]
        public int ExpiresIn { get; init; }

        [JsonPropertyName("refresh_token")]
        public string? RefreshToken { get; init; }
    }
}
