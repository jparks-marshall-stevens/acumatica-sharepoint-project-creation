using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.Acumatica;

/// <summary>
/// Reads projects from an Acumatica Generic Inquiry exposed over the OData (v4) feed.
/// Server-side <c>$filter</c> restricts results to those created after the last-run watermark.
/// </summary>
public sealed class AcumaticaClient : IAcumaticaClient
{
    private readonly HttpClient _http;
    private readonly AcumaticaTokenProvider _tokenProvider;
    private readonly AcumaticaOptions _options;
    private readonly ILogger<AcumaticaClient> _logger;

    public AcumaticaClient(
        HttpClient http,
        AcumaticaTokenProvider tokenProvider,
        IOptions<AcumaticaOptions> options,
        ILogger<AcumaticaClient> logger)
    {
        _http = http;
        _tokenProvider = tokenProvider;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<IReadOnlyList<AcumaticaProject>> GetProjectsCreatedAfterAsync(
        DateTimeOffset createdAfterUtc,
        CancellationToken cancellationToken)
    {
        var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);

        // OData v4 GI feed. Filter on the created date/time column; order oldest-first so
        // processing failures leave the watermark on the earliest unprocessed record.
        var createdField = _options.CreatedDateTimeField;
        var filterValue = createdAfterUtc.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

        var baseUrl = _options.BaseUrl.TrimEnd('/');
        var url = $"{baseUrl}/t/{Uri.EscapeDataString(_options.Tenant)}/api/odata/gi/" +
                  $"{Uri.EscapeDataString(_options.GenericInquiryName)}" +
                  $"?$filter={Uri.EscapeDataString($"{createdField} gt {filterValue}")}" +
                  $"&$orderby={Uri.EscapeDataString($"{createdField} asc")}";

        _logger.LogInformation("Querying Acumatica GI '{Gi}' for projects created after {After:o}",
            _options.GenericInquiryName, createdAfterUtc);

        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new InvalidOperationException(
                $"Acumatica GI query failed ({(int)response.StatusCode}): {body}");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);

        if (!doc.RootElement.TryGetProperty("value", out var valueArray) ||
            valueArray.ValueKind != JsonValueKind.Array)
        {
            _logger.LogWarning("Acumatica GI response did not contain a 'value' array; returning empty set.");
            return Array.Empty<AcumaticaProject>();
        }

        var results = new List<AcumaticaProject>();
        foreach (var row in valueArray.EnumerateArray())
        {
            var projectId = GetString(row, _options.ProjectIdField);
            if (string.IsNullOrWhiteSpace(projectId))
            {
                _logger.LogWarning("Skipping GI row with no value in ProjectId field '{Field}'.", _options.ProjectIdField);
                continue;
            }

            results.Add(new AcumaticaProject
            {
                ProjectId = projectId!.Trim(),
                ProjectName = GetString(row, _options.ProjectNameField),
                CustomerName = GetString(row, _options.CustomerNameField),
                ProjectManager = GetString(row, _options.ProjectManagerField),
                Practice = GetString(row, _options.PracticeField),
                CreatedDateTime = GetDateTime(row, _options.CreatedDateTimeField),
            });
        }

        _logger.LogInformation("Acumatica GI returned {Count} project(s).", results.Count);
        return results;
    }

    private static string? GetString(JsonElement row, string property)
    {
        if (string.IsNullOrEmpty(property) || !row.TryGetProperty(property, out var el))
        {
            return null;
        }

        return el.ValueKind switch
        {
            JsonValueKind.String => el.GetString(),
            JsonValueKind.Number => el.GetRawText(),
            JsonValueKind.True or JsonValueKind.False => el.GetBoolean().ToString(),
            JsonValueKind.Null or JsonValueKind.Undefined => null,
            _ => el.GetRawText(),
        };
    }

    private static DateTimeOffset? GetDateTime(JsonElement row, string property)
    {
        if (string.IsNullOrEmpty(property) || !row.TryGetProperty(property, out var el) ||
            el.ValueKind != JsonValueKind.String)
        {
            return null;
        }

        return DateTimeOffset.TryParse(el.GetString(), CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dto)
            ? dto
            : null;
    }
}
