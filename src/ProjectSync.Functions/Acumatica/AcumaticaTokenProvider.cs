using System.Net.Http.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.Acumatica;

/// <summary>
/// Acquires and caches an OAuth2 bearer token from Acumatica's identity endpoint
/// using the client-credentials grant. Tokens are refreshed slightly before expiry.
/// </summary>
public sealed class AcumaticaTokenProvider
{
    private readonly HttpClient _http;
    private readonly AcumaticaOptions _options;
    private readonly ILogger<AcumaticaTokenProvider> _logger;
    private readonly SemaphoreSlim _lock = new(1, 1);

    private string? _cachedToken;
    private DateTimeOffset _expiresAtUtc = DateTimeOffset.MinValue;

    public AcumaticaTokenProvider(
        HttpClient http,
        IOptions<AcumaticaOptions> options,
        ILogger<AcumaticaTokenProvider> logger)
    {
        _http = http;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
    {
        // Fast path: still valid (with 60s safety margin).
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

            var tokenUrl = $"{_options.BaseUrl.TrimEnd('/')}/identity/connect/token";
            var form = new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = _options.ClientId,
                ["client_secret"] = _options.ClientSecret,
                ["scope"] = _options.Scope,
            };

            _logger.LogDebug("Requesting Acumatica OAuth token from {TokenUrl}", tokenUrl);

            using var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl)
            {
                Content = new FormUrlEncodedContent(form),
            };

            using var response = await _http.SendAsync(request, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                throw new InvalidOperationException(
                    $"Acumatica token request failed ({(int)response.StatusCode}): {body}");
            }

            var token = await response.Content.ReadFromJsonAsync<TokenResponse>(cancellationToken)
                ?? throw new InvalidOperationException("Acumatica token response was empty.");

            _cachedToken = token.AccessToken;
            _expiresAtUtc = DateTimeOffset.UtcNow.AddSeconds(token.ExpiresIn);
            _logger.LogDebug("Acquired Acumatica token, expires in {Seconds}s", token.ExpiresIn);
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

        [JsonPropertyName("token_type")]
        public string TokenType { get; init; } = "Bearer";
    }
}
