using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using ProjectSync.Options;

namespace ProjectSync.SharePoint;

/// <summary>
/// Creates an anonymous, upload-only ("Request files") sharing link on a folder via Microsoft Graph.
/// Uses the same app-only certificate as the CSOM path, but Graph requires its own application grant
/// (Graph <c>Sites.Selected</c> on the target site, or <c>Files.ReadWrite.All</c>). Everything here is
/// fail-soft: any failure returns <c>null</c> so document-set creation still succeeds.
/// </summary>
public sealed class GraphUploadLinkService
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";

    // One shared client; auth is applied per-request (never on DefaultRequestHeaders) so concurrent
    // document-set creations don't race on shared state.
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };

    private readonly SharePointContextFactory _contextFactory;
    private readonly SharePointOptions _options;
    private readonly ILogger<GraphUploadLinkService> _logger;
    private readonly Lazy<IConfidentialClientApplication> _app;

    public GraphUploadLinkService(
        SharePointContextFactory contextFactory,
        IOptions<SharePointOptions> options,
        ILogger<GraphUploadLinkService> logger)
    {
        _contextFactory = contextFactory;
        _options = options.Value;
        _logger = logger;
        _app = new Lazy<IConfidentialClientApplication>(() =>
            ConfidentialClientApplicationBuilder.Create(_options.ClientId)
                .WithCertificate(_contextFactory.Certificate)
                .WithTenantId(_options.AzureAdTenant)
                .Build());
    }

    /// <summary>
    /// Mints an upload-only sharing link for the folder at <paramref name="folderServerRelativeUrl"/>.
    /// Returns the link URL, or <c>null</c> if anything fails (logged as a warning).
    /// </summary>
    /// <param name="siteUrl">Absolute site URL, e.g. https://contoso.sharepoint.com/sites/GiftEstate.</param>
    /// <param name="listRootServerRelativeUrl">Server-relative URL of the library root (from CSOM).</param>
    /// <param name="folderServerRelativeUrl">Server-relative URL of the folder to share.</param>
    public async Task<string?> CreateUploadLinkAsync(
        string siteUrl,
        string listRootServerRelativeUrl,
        string folderServerRelativeUrl,
        CancellationToken cancellationToken)
    {
        try
        {
            var token = await AcquireTokenAsync(cancellationToken);

            var siteId = await ResolveSiteIdAsync(token, siteUrl, cancellationToken);
            if (siteId is null)
            {
                return null;
            }

            var driveId = await ResolveDriveIdAsync(token, siteId, listRootServerRelativeUrl, cancellationToken);
            if (driveId is null)
            {
                return null;
            }

            var relativePath = folderServerRelativeUrl
                .Substring(Math.Min(listRootServerRelativeUrl.Length, folderServerRelativeUrl.Length))
                .Trim('/');
            var encodedPath = string.Join('/', relativePath
                .Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(Uri.EscapeDataString));

            // Resolve the folder to an explicit driveItem id. Path addressing (root:/path:) on createLink
            // has proven unreliable here (links bound to the drive root), so we bind by item id instead.
            var itemUrl = $"{GraphBase}/drives/{driveId}/root:/{encodedPath}?$select=id,name,parentReference";
            var (okItem, itemJson) = await GetAsync(token, itemUrl, cancellationToken);
            if (!okItem)
            {
                _logger.LogWarning("Graph item lookup failed for '{Folder}': {Body}", folderServerRelativeUrl, Truncate(itemJson));
                return null;
            }

            using var itemDoc = JsonDocument.Parse(itemJson);
            var itemId = itemDoc.RootElement.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            var itemName = itemDoc.RootElement.TryGetProperty("name", out var nameEl) ? nameEl.GetString() : "?";
            if (string.IsNullOrEmpty(itemId))
            {
                _logger.LogWarning("Graph item lookup returned no id for '{Folder}'.", folderServerRelativeUrl);
                return null;
            }

            _logger.LogInformation(
                "Upload-link target resolved to item '{Name}' (id {Id}) via path '{Path}'.", itemName, itemId, encodedPath);

            var expiration = DateTimeOffset.UtcNow.AddDays(_options.ClientUploadLinkExpirationDays);
            var body = new
            {
                type = "createOnly",
                scope = _options.ClientUploadLinkScope,
                expirationDateTime = expiration.ToString("yyyy-MM-ddTHH:mm:ssZ"),
            };

            var url = $"{GraphBase}/drives/{driveId}/items/{itemId}/createLink";
            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"),
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            using var response = await Http.SendAsync(request, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning(
                    "Graph createLink failed ({Status}) for '{Folder}': {Body}",
                    (int)response.StatusCode, folderServerRelativeUrl, Truncate(json));
                return null;
            }

            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("link", out var link) &&
                link.TryGetProperty("webUrl", out var webUrl))
            {
                var result = webUrl.GetString();
                _logger.LogInformation(
                    "Created {Scope} upload link (expires {Expiry:yyyy-MM-dd}) for '{Folder}'.",
                    _options.ClientUploadLinkScope, expiration, folderServerRelativeUrl);
                return result;
            }

            _logger.LogWarning("Graph createLink returned no link.webUrl for '{Folder}'.", folderServerRelativeUrl);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to create upload link for '{Folder}'.", folderServerRelativeUrl);
            return null;
        }
    }

    private async Task<string> AcquireTokenAsync(CancellationToken cancellationToken)
    {
        var result = await _app.Value
            .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
            .ExecuteAsync(cancellationToken);
        return result.AccessToken;
    }

    /// <summary>Resolves the Graph site id from a site URL via the <c>{host}:{path}</c> addressing form.</summary>
    private async Task<string?> ResolveSiteIdAsync(string token, string siteUrl, CancellationToken cancellationToken)
    {
        var uri = new Uri(siteUrl);
        var path = uri.AbsolutePath.Trim('/');
        var url = $"{GraphBase}/sites/{uri.Host}:/{path}";
        var (ok, json) = await GetAsync(token, url, cancellationToken);
        if (!ok)
        {
            _logger.LogWarning("Graph site lookup failed for '{Site}': {Body}", siteUrl, Truncate(json));
            return null;
        }

        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("id", out var id) ? id.GetString() : null;
    }

    /// <summary>
    /// Finds the drive backing the library whose root folder is <paramref name="listRootServerRelativeUrl"/>,
    /// by matching each drive's <c>webUrl</c> path. Falls back to the site's default drive.
    /// </summary>
    private async Task<string?> ResolveDriveIdAsync(
        string token, string siteId, string listRootServerRelativeUrl, CancellationToken cancellationToken)
    {
        var url = $"{GraphBase}/sites/{siteId}/drives?$select=id,name,webUrl";
        var (ok, json) = await GetAsync(token, url, cancellationToken);
        if (!ok)
        {
            _logger.LogWarning("Graph drives lookup failed: {Body}", Truncate(json));
            return null;
        }

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("value", out var drives) || drives.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var target = listRootServerRelativeUrl.Trim('/');
        string? firstId = null;
        foreach (var drive in drives.EnumerateArray())
        {
            var id = drive.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            firstId ??= id;
            if (drive.TryGetProperty("webUrl", out var webUrlEl) &&
                Uri.TryCreate(webUrlEl.GetString(), UriKind.Absolute, out var driveUri) &&
                string.Equals(Uri.UnescapeDataString(driveUri.AbsolutePath).Trim('/'), target, StringComparison.OrdinalIgnoreCase))
            {
                return id;
            }
        }

        if (firstId is not null)
        {
            _logger.LogWarning(
                "No Graph drive matched library root '{Root}'; falling back to the site's first drive.",
                listRootServerRelativeUrl);
        }

        return firstId;
    }

    private static async Task<(bool Ok, string Body)> GetAsync(string token, string url, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        using var response = await Http.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        return (response.IsSuccessStatusCode, body);
    }

    private static string Truncate(string value) => value.Length <= 500 ? value : value[..500];
}
