using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.HubSpot;
using ProjectSync.Options;

// -----------------------------------------------------------------------------
// HubSpot connectivity + discovery test.
//
// Confirms the private-app token works and DISCOVERS the values you need to
// configure the scoping sync:
//   • deal pipelines + their stages (ids → labels)  → HubSpot:PipelineId / StageIds
//   • candidate deal properties (customer/practice)  → HubSpot:CustomerProperty / PracticeProperty
//   • a sample of recently-modified deals via the typed client
//
// Config comes from the Functions local.settings.json "Values" section (git-ignored),
// env vars override (HubSpot__AccessToken=...). Never commit the token.
// -----------------------------------------------------------------------------

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var options = new HubSpotOptions();
configuration.GetSection(HubSpotOptions.SectionName).Bind(options);

if (string.IsNullOrWhiteSpace(options.RefreshToken) && string.IsNullOrWhiteSpace(options.AccessToken))
{
    Console.Error.WriteLine("❌ No HubSpot credentials configured.");
    Console.Error.WriteLine($"   Add to the 'Values' in {localSettingsPath}:");
    Console.Error.WriteLine("     \"HubSpot:ClientId\", \"HubSpot:ClientSecret\", and \"HubSpot:RefreshToken\"");
    Console.Error.WriteLine("   Run tools/HubSpotOAuthSetup first to obtain the refresh token.");
    return 1;
}

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Information)
    .AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; }));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(options.TimeoutSeconds) };
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
var baseUrl = options.BaseUrl.TrimEnd('/');

var tokenProvider = new HubSpotTokenProvider(http, Options.Create(options), loggerFactory.CreateLogger<HubSpotTokenProvider>());
string accessToken;
try
{
    accessToken = await tokenProvider.GetAccessTokenAsync(CancellationToken.None);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Could not obtain a HubSpot access token: {ex.Message}");
    return 1;
}
http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

Console.WriteLine("=== HubSpot connectivity + discovery test ===");
Console.WriteLine($"Base URL : {baseUrl}");
Console.WriteLine($"Auth     : {(string.IsNullOrWhiteSpace(options.RefreshToken) ? "static token" : "OAuth refresh token")} → access token acquired ({accessToken.Length} chars)");
Console.WriteLine();

// --- Step 1: pipelines + stages -------------------------------------------
try
{
    Console.WriteLine("→ Deal pipelines and stages (use these ids for HubSpot:PipelineId / StageIds):");
    using var resp = await http.GetAsync($"{baseUrl}/crm/v3/pipelines/deals");
    var body = await resp.Content.ReadAsStringAsync();
    if (!resp.IsSuccessStatusCode)
    {
        Console.Error.WriteLine($"❌ pipelines failed ({(int)resp.StatusCode}): {Truncate(body, 300)}");
        if ((int)resp.StatusCode == 401) Console.Error.WriteLine("   → token invalid or missing 'crm.objects.deals.read' scope.");
        return 2;
    }

    using var doc = JsonDocument.Parse(body);
    foreach (var pl in doc.RootElement.GetProperty("results").EnumerateArray())
    {
        Console.WriteLine($"  Pipeline  \"{pl.GetProperty("label").GetString()}\"  id={pl.GetProperty("id").GetString()}");
        foreach (var st in pl.GetProperty("stages").EnumerateArray())
        {
            Console.WriteLine($"      stage  \"{st.GetProperty("label").GetString(),-28}\"  id={st.GetProperty("id").GetString()}");
        }
    }
    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ pipelines probe failed: {ex.Message}");
    return 2;
}

// --- Step 2: candidate properties -----------------------------------------
try
{
    Console.WriteLine("→ Candidate deal properties (name → label) for customer/practice mapping:");
    using var resp = await http.GetAsync($"{baseUrl}/crm/v3/properties/deals");
    var body = await resp.Content.ReadAsStringAsync();
    if (resp.IsSuccessStatusCode)
    {
        using var doc = JsonDocument.Parse(body);
        var all = doc.RootElement.GetProperty("results");
        var keywords = new[] { "customer", "company", "account", "practice", "service", "client", "project" };
        var matches = all.EnumerateArray()
            .Select(p => (name: p.GetProperty("name").GetString() ?? "", label: p.GetProperty("label").GetString() ?? ""))
            .Where(p => keywords.Any(k =>
                p.name.Contains(k, StringComparison.OrdinalIgnoreCase) ||
                p.label.Contains(k, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(p => p.name)
            .ToList();

        Console.WriteLine($"  ({all.GetArrayLength()} total properties; {matches.Count} match customer/practice keywords)");
        foreach (var (name, label) in matches)
        {
            Console.WriteLine($"      {name,-32} → {label}");
        }
    }
    else
    {
        Console.Error.WriteLine($"⚠ properties probe failed ({(int)resp.StatusCode}): {Truncate(body, 200)}");
    }
    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"⚠ properties probe failed: {ex.Message}");
}

// --- Step 3: typed client sample ------------------------------------------
try
{
    var client = new HubSpotClient(http, tokenProvider, Options.Create(options), loggerFactory.CreateLogger<HubSpotClient>());
    var since = DateTimeOffset.UtcNow.AddDays(-90);
    Console.WriteLine($"→ Typed client: deals modified after {since:yyyy-MM-dd}" +
        (options.PipelineIds.Count == 0 ? " (all pipelines)" : $" ({options.PipelineIds.Count} pipeline(s))") + " (sample of 25)…");
    var deals = await client.GetDealsModifiedAfterAsync(since, maxResults: 25, CancellationToken.None);
    Console.WriteLine($"✔ Mapped {deals.Count} deal(s).");
    Console.WriteLine();

    foreach (var d in deals.OrderByDescending(x => x.ModifiedAt).Take(8))
    {
        Console.WriteLine($"    [{d.DealId}] {Truncate(d.DealName, 32)} | cust={Truncate(d.CustomerName, 20) ?? "<none>"} | " +
            $"practice={d.Practice ?? "<none>"} | stage={d.StageId} | modified={d.ModifiedAt:yyyy-MM-dd}");
    }

    if (deals.Count > 0)
    {
        Console.WriteLine();
        Console.WriteLine("  Raw properties on the first deal (to confirm mapping):");
        foreach (var kv in deals[0].Properties.OrderBy(k => k.Key))
        {
            Console.WriteLine($"      {kv.Key,-32} → {Truncate(kv.Value, 48)}");
        }
    }
    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ typed client failed: {ex.Message}");
    return 3;
}

Console.WriteLine("✅ HubSpot connectivity OK.");
return 0;

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!File.Exists(path)) return result;
    using var doc = JsonDocument.Parse(File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var values))
    {
        foreach (var prop in values.EnumerateObject())
        {
            result[prop.Name] = prop.Value.GetString();
        }
    }
    return result;
}

static string? Truncate(string? s, int max)
    => string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s[..max] + "…");
