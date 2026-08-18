using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.HubSpot;

/// <summary>
/// Reads deals from HubSpot via the CRM v3 search API using a private-app access token.
/// </summary>
public sealed class HubSpotClient : IHubSpotClient
{
    // Safety cap on pages (100 deals/page) so a mis-set watermark can't loop unbounded.
    private const int MaxPages = 200;

    // HubSpot's search API returns 400 when paging past this offset.
    private const int SearchWindowLimit = 10000;

    private readonly HttpClient _http;
    private readonly HubSpotTokenProvider _tokenProvider;
    private readonly HubSpotOptions _options;
    private readonly ILogger<HubSpotClient> _logger;

    public HubSpotClient(
        HttpClient http,
        HubSpotTokenProvider tokenProvider,
        IOptions<HubSpotOptions> options,
        ILogger<HubSpotClient> logger)
    {
        _http = http;
        _tokenProvider = tokenProvider;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<IReadOnlyList<HubSpotDeal>> GetDealsModifiedAfterAsync(
        DateTimeOffset modifiedAfterUtc, int maxResults, CancellationToken cancellationToken)
    {
        var url = $"{_options.BaseUrl.TrimEnd('/')}/crm/v3/objects/deals/search";
        var properties = BuildRequestedProperties();
        var sinceMs = modifiedAfterUtc.ToUnixTimeMilliseconds();

        var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);
        var deals = new List<HubSpotDeal>();
        string? after = null;
        for (var page = 0; page < MaxPages; page++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var body = BuildSearchBody(sinceMs, properties, after);
            var json = await PostSearchAsync(url, body, token, cancellationToken);

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("results", out var results) && results.ValueKind == JsonValueKind.Array)
            {
                foreach (var result in results.EnumerateArray())
                {
                    deals.Add(MapDeal(result));
                    if (deals.Count >= maxResults)
                    {
                        _logger.LogInformation("HubSpot: reached maxResults={Max}; stopping.", maxResults);
                        return deals;
                    }
                }
            }

            after = root.TryGetProperty("paging", out var paging) &&
                    paging.TryGetProperty("next", out var next) &&
                    next.TryGetProperty("after", out var afterEl)
                ? afterEl.GetString()
                : null;

            if (string.IsNullOrEmpty(after))
            {
                break;
            }

            // HubSpot search cannot page beyond 10,000 results. Stop at the wall; the ascending
            // modified-date sort + watermark means the remainder is picked up on the next poll.
            if (int.TryParse(after, out var offset) && offset >= SearchWindowLimit)
            {
                _logger.LogWarning(
                    "HubSpot: hit the {Limit}-result search window ({Count} so far); remaining deals continue next cycle.",
                    SearchWindowLimit, deals.Count);
                break;
            }
        }

        _logger.LogInformation("HubSpot returned {Count} deal(s) modified after {Since:o}.", deals.Count, modifiedAfterUtc);
        return deals;
    }

    /// <summary>
    /// POSTs a search request, retrying on HTTP 429 (HubSpot's search API is rate-limited to a few
    /// requests/second). Honors the Retry-After header when present, else uses a short backoff.
    /// </summary>
    private async Task<string> PostSearchAsync(string url, string body, string token, CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        for (var attempt = 1; ; attempt++)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            using var response = await _http.SendAsync(request, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);
            if (response.IsSuccessStatusCode)
            {
                return json;
            }

            if ((int)response.StatusCode == 429 && attempt < maxAttempts)
            {
                var delay = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromMilliseconds(500 * attempt);
                _logger.LogWarning("HubSpot search rate-limited (429); retry {Attempt}/{Max} after {Delay}.",
                    attempt, maxAttempts, delay);
                await Task.Delay(delay, cancellationToken);
                continue;
            }

            throw new HttpRequestException(
                $"HubSpot deals/search failed ({(int)response.StatusCode}): {Truncate(json, 500)}");
        }
    }

    public async Task<string?> ResolveCustomerNameAsync(HubSpotDeal deal, CancellationToken cancellationToken)
    {
        try
        {
            var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);
            var contactId = deal.ClientContactId;
            if (string.IsNullOrWhiteSpace(contactId))
            {
                contactId = await FindClientContactIdAsync(deal.DealId, token, cancellationToken);
            }

            if (!string.IsNullOrWhiteSpace(contactId))
            {
                var contact = await GetJsonAsync(
                    $"{Base()}/crm/v3/objects/contacts/{contactId}?properties=company,associatedcompanyid", token, cancellationToken);
                if (contact is { } c && c.RootElement.TryGetProperty("properties", out var cp))
                {
                    var companyText = cp.TryGetProperty("company", out var ctEl) ? ctEl.GetString() : null;
                    var assocId = cp.TryGetProperty("associatedcompanyid", out var acEl) ? acEl.GetString() : null;
                    var assocName = await GetCompanyNameAsync(assocId, token, cancellationToken);

                    // Order per config: text-first or associated-company-first; deal name is the fallback below.
                    var first = _options.CustomerCompanyTextFirst ? companyText : assocName;
                    var second = _options.CustomerCompanyTextFirst ? assocName : companyText;
                    if (!string.IsNullOrWhiteSpace(first)) return first;
                    if (!string.IsNullOrWhiteSpace(second)) return second;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not resolve customer for deal {DealId}; falling back to deal name.", deal.DealId);
        }

        return deal.DealName; // fallback: no client contact / no company
    }

    public async Task<IReadOnlyDictionary<string, string>> GetOwnerEmailsAsync(CancellationToken cancellationToken)
    {
        var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var url = $"{Base()}/crm/v3/owners?limit=100";
        for (var page = 0; page < 50 && url is not null; page++)
        {
            var doc = await GetJsonAsync(url, token, cancellationToken);
            if (doc is null) break;
            var root = doc.RootElement;
            if (root.TryGetProperty("results", out var results))
            {
                foreach (var o in results.EnumerateArray())
                {
                    var id = o.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
                    var email = o.TryGetProperty("email", out var emEl) ? emEl.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(email))
                    {
                        map[id!] = email!;
                    }
                }
            }

            url = root.TryGetProperty("paging", out var p) && p.TryGetProperty("next", out var nx) && nx.TryGetProperty("link", out var lk)
                ? lk.GetString()
                : null;
        }

        return map;
    }

    /// <summary>Resolves a company id to its name (null if the id is blank or the lookup fails).</summary>
    private async Task<string?> GetCompanyNameAsync(string? companyId, string token, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(companyId))
        {
            return null;
        }

        var doc = await GetJsonAsync($"{Base()}/crm/v3/objects/companies/{companyId}?properties=name", token, cancellationToken);
        return doc is not null &&
               doc.RootElement.TryGetProperty("properties", out var props) &&
               props.TryGetProperty("name", out var name)
            ? name.GetString()
            : null;
    }

    /// <summary>Finds the client contact id via the labeled deal→contact association (v4).</summary>
    private async Task<string?> FindClientContactIdAsync(string dealId, string token, CancellationToken cancellationToken)
    {
        var doc = await GetJsonAsync($"{Base()}/crm/v4/objects/deals/{dealId}/associations/contacts", token, cancellationToken);
        if (doc is null || !doc.RootElement.TryGetProperty("results", out var results))
        {
            return null;
        }

        foreach (var r in results.EnumerateArray())
        {
            if (r.TryGetProperty("associationTypes", out var types))
            {
                var isClient = types.EnumerateArray().Any(t =>
                    t.TryGetProperty("label", out var lbl) &&
                    string.Equals(lbl.GetString(), _options.ClientContactLabel, StringComparison.OrdinalIgnoreCase));
                if (isClient)
                {
                    return r.TryGetProperty("toObjectId", out var to) ? to.GetRawText().Trim('"') : null;
                }
            }
        }

        return null;
    }

    public async Task<IReadOnlyList<HubSpotDeal>> GetDealsByIdAsync(
        IReadOnlyList<string> dealIds, CancellationToken cancellationToken)
    {
        const int batchSize = 100; // HubSpot's batch-read limit.
        var wanted = dealIds
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (wanted.Count == 0)
        {
            return Array.Empty<HubSpotDeal>();
        }

        var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);
        var properties = BuildRequestedProperties();
        var results = new List<HubSpotDeal>(wanted.Count);

        for (var offset = 0; offset < wanted.Count; offset += batchSize)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var chunk = wanted.Skip(offset).Take(batchSize).ToList();
            var payload = new Dictionary<string, object?>
            {
                ["properties"] = properties,
                ["inputs"] = chunk.Select(id => new Dictionary<string, string> { ["id"] = id }).ToList(),
            };

            // Reuses the search POST helper for its 429 handling; the batch endpoint is rate-limited too.
            var json = await PostSearchAsync(
                $"{Base()}/crm/v3/objects/deals/batch/read",
                JsonSerializer.Serialize(payload),
                token,
                cancellationToken);

            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("results", out var arr) || arr.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var result in arr.EnumerateArray())
            {
                results.Add(MapDeal(result));
            }
        }

        var missing = wanted.Count - results.Count;
        if (missing > 0)
        {
            _logger.LogWarning("HubSpot batch read: {Missing} of {Requested} deal id(s) returned nothing.",
                missing, wanted.Count);
        }

        return results;
    }

    private string Base() => _options.BaseUrl.TrimEnd('/');

    private async Task<JsonDocument?> GetJsonAsync(string url, string token, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        using var response = await _http.SendAsync(request, cancellationToken);
        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogWarning("HubSpot GET {Url} failed ({Status}): {Body}", url, (int)response.StatusCode, Truncate(json, 200));
            return null;
        }

        return JsonDocument.Parse(json);
    }

    private List<string> BuildRequestedProperties()
    {
        var props = new List<string>
        {
            _options.DealNameProperty,
            _options.OwnerIdProperty,
            _options.CreatedProperty,
            _options.ModifiedProperty,
            "dealstage",
            "pipeline",
            "hs_object_id",
        };
        if (!string.IsNullOrWhiteSpace(_options.CustomerProperty)) props.Add(_options.CustomerProperty);
        if (!string.IsNullOrWhiteSpace(_options.PracticeProperty)) props.Add(_options.PracticeProperty);
        if (!string.IsNullOrWhiteSpace(_options.ClientContactIdProperty)) props.Add(_options.ClientContactIdProperty);
        if (!string.IsNullOrWhiteSpace(_options.OpportunityIdProperty)) props.Add(_options.OpportunityIdProperty);
        props.AddRange(_options.ExtraProperties);

        return props
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private string BuildSearchBody(long sinceMs, List<string> properties, string? after)
    {
        // Filters are dictionaries (not anonymous objects in a List<object>) so System.Text.Json
        // serializes their contents — a List<object> would serialize each element as an empty {}.
        var filters = new List<Dictionary<string, object>>
        {
            new() { ["propertyName"] = _options.ModifiedProperty, ["operator"] = "GT", ["value"] = sinceMs.ToString() },
        };
        if (_options.PipelineIds.Count > 0)
        {
            filters.Add(new() { ["propertyName"] = "pipeline", ["operator"] = "IN", ["values"] = _options.PipelineIds });
        }
        if (_options.TerminalStageIds.Count > 0)
        {
            // "In scoping" = not yet Won/Lost/Closed.
            filters.Add(new() { ["propertyName"] = "dealstage", ["operator"] = "NOT_IN", ["values"] = _options.TerminalStageIds });
        }

        var payload = new Dictionary<string, object?>
        {
            ["filterGroups"] = new[] { new Dictionary<string, object> { ["filters"] = filters } },
            ["sorts"] = new[] { new Dictionary<string, object> { ["propertyName"] = _options.ModifiedProperty, ["direction"] = "ASCENDING" } },
            ["properties"] = properties,
            ["limit"] = 100,
        };
        if (!string.IsNullOrEmpty(after))
        {
            payload["after"] = after;
        }

        return JsonSerializer.Serialize(payload);
    }

    private HubSpotDeal MapDeal(JsonElement result)
    {
        var id = result.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;

        var props = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        if (result.TryGetProperty("properties", out var propsEl) && propsEl.ValueKind == JsonValueKind.Object)
        {
            foreach (var p in propsEl.EnumerateObject())
            {
                props[p.Name] = p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString() : p.Value.GetRawText();
            }
        }

        string? Get(string key) => !string.IsNullOrWhiteSpace(key) && props.TryGetValue(key, out var v) ? v : null;

        return new HubSpotDeal
        {
            DealId = id ?? Get("hs_object_id") ?? string.Empty,
            DealName = Get(_options.DealNameProperty),
            CustomerName = Get(_options.CustomerProperty),
            Practice = Get(_options.PracticeProperty),
            StageId = Get("dealstage"),
            PipelineId = Get("pipeline"),
            OwnerId = Get(_options.OwnerIdProperty),
            ClientContactId = Get(_options.ClientContactIdProperty),
            OpportunityId = Get(_options.OpportunityIdProperty)?.Trim(),
            CreatedAt = ParseTop(result, "createdAt"),
            ModifiedAt = ParseTop(result, "updatedAt"),
            Properties = props,
        };
    }

    private static DateTimeOffset? ParseTop(JsonElement result, string name)
        => result.TryGetProperty(name, out var el) &&
           el.ValueKind == JsonValueKind.String &&
           DateTimeOffset.TryParse(el.GetString(), out var dto)
            ? dto
            : null;

    private static string Truncate(string? s, int max)
        => string.IsNullOrEmpty(s) ? string.Empty : (s.Length <= max ? s : s[..max] + "…");
}
