using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Acumatica;
using ProjectSync.Options;

// -----------------------------------------------------------------------------
// Acumatica connectivity test.
//
// Exercises the same code the Azure Function uses (OAuth token + GI OData query)
// without needing the Functions runtime, Azurite, or SharePoint.
//
// Configuration is read from the Functions local.settings.json "Values" section,
// with environment variables overriding (use double underscore, e.g.
// Acumatica__ClientSecret=...). Put real credentials in local.settings.json
// (it is git-ignored) or in env vars — never commit secrets.
// -----------------------------------------------------------------------------

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var inMemory = LoadFunctionsValues(localSettingsPath);

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(inMemory)
    .AddEnvironmentVariables()
    .Build();

var options = new AcumaticaOptions();
configuration.GetSection(AcumaticaOptions.SectionName).Bind(options);

if (!Validate(options, out var problem))
{
    Console.Error.WriteLine($"❌ Configuration incomplete: {problem}");
    Console.Error.WriteLine();
    Console.Error.WriteLine($"Edit the 'Values' in {localSettingsPath}");
    Console.Error.WriteLine("and set real values for:");
    Console.Error.WriteLine("  Acumatica:BaseUrl, Acumatica:Tenant, Acumatica:ClientId, Acumatica:ClientSecret");
    Console.Error.WriteLine("Or set them as env vars (Acumatica__BaseUrl, Acumatica__ClientSecret, ...).");
    return 1;
}

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Debug)
    .AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; }));

var optionsWrapper = Microsoft.Extensions.Options.Options.Create(options);
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(options.TimeoutSeconds) };

Console.WriteLine("=== Acumatica connectivity test ===");
Console.WriteLine($"Base URL : {options.BaseUrl}");
Console.WriteLine($"Tenant   : {options.Tenant}");
Console.WriteLine($"GI       : {options.GenericInquiryName}");
Console.WriteLine();

// --- Step 1: OAuth token ---------------------------------------------------
var tokenProvider = new AcumaticaTokenProvider(
    http, optionsWrapper, loggerFactory.CreateLogger<AcumaticaTokenProvider>());

string token;
try
{
    Console.WriteLine("→ Requesting OAuth token…");
    token = await tokenProvider.GetAccessTokenAsync(CancellationToken.None);
    Console.WriteLine($"✔ Token acquired ({token.Length} chars).");
    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Token request failed: {ex.Message}");
    return 2;
}

// --- Step 2: raw first row (confirms exact OData property names) ------------
try
{
    var baseUrl = options.BaseUrl.TrimEnd('/');
    var rawUrl = $"{baseUrl}/t/{Uri.EscapeDataString(options.Tenant)}/api/odata/gi/" +
                 $"{Uri.EscapeDataString(options.GenericInquiryName)}?$top=1";

    Console.WriteLine("→ Fetching one raw row to confirm property names…");
    using var req = new HttpRequestMessage(HttpMethod.Get, rawUrl);
    req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

    using var resp = await http.SendAsync(req, CancellationToken.None);
    var body = await resp.Content.ReadAsStringAsync();
    if (!resp.IsSuccessStatusCode)
    {
        Console.Error.WriteLine($"❌ GI query failed ({(int)resp.StatusCode}): {body}");
        return 3;
    }

    using var doc = JsonDocument.Parse(body);
    if (doc.RootElement.TryGetProperty("value", out var arr) &&
        arr.ValueKind == JsonValueKind.Array && arr.GetArrayLength() > 0)
    {
        var first = arr[0];
        Console.WriteLine("✔ Available properties in the GI feed (name → sample value):");
        foreach (var prop in first.EnumerateObject())
        {
            var val = prop.Value.ValueKind == JsonValueKind.String
                ? prop.Value.GetString()
                : prop.Value.GetRawText();
            Console.WriteLine($"    {prop.Name,-28} → {Truncate(val, 50)}");
        }

        Console.WriteLine();
        Console.WriteLine("  Confirm these configured field names exist above:");
        foreach (var (label, field) in new[]
        {
            ("ProjectIdField", options.ProjectIdField),
            ("CustomerNameField", options.CustomerNameField),
            ("ProjectNameField", options.ProjectNameField),
            ("ProjectManagerField", options.ProjectManagerField),
            ("PracticeField", options.PracticeField),
            ("CreatedDateTimeField", options.CreatedDateTimeField),
        })
        {
            var present = first.TryGetProperty(field, out _);
            Console.WriteLine($"    {(present ? "✔" : "✖")} {label,-22} = '{field}'");
        }
    }
    else
    {
        Console.WriteLine("⚠ GI returned no rows (empty inquiry). Property-name check skipped.");
    }

    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Raw GI fetch failed: {ex.Message}");
    return 3;
}

// --- Step 3: typed client (the exact function path) ------------------------
try
{
    var client = new AcumaticaClient(
        http, tokenProvider, optionsWrapper, loggerFactory.CreateLogger<AcumaticaClient>());

    // Wide window for the connectivity test so we actually see mapped rows regardless of how old
    // the sample data is. (The function itself uses a 15-min watermark, not this.)
    var since = DateTimeOffset.UtcNow.AddYears(-5);
    Console.WriteLine($"→ Running the typed client: projects created after {since:o}…");
    var projects = await client.GetProjectsCreatedAfterAsync(since, CancellationToken.None);

    Console.WriteLine($"✔ Mapped {projects.Count} project(s).");
    Console.WriteLine();

    // Practice distribution — what values actually exist, and how many per practice.
    Console.WriteLine("Practice breakdown (value → count):");
    var byPractice = projects
        .GroupBy(p => string.IsNullOrWhiteSpace(p.Practice) ? "<blank>" : p.Practice!)
        .OrderByDescending(g => g.Count());
    foreach (var g in byPractice)
    {
        Console.WriteLine($"    {g.Key,-28} → {g.Count()}");
    }
    Console.WriteLine();

    // Real Estate & Gift examples (the practice we intend to sync).
    var eg = projects
        .Where(p => string.Equals(p.Practice?.Trim(), "Estate & Gift", StringComparison.OrdinalIgnoreCase))
        .OrderByDescending(p => p.CreatedDateTime)
        .Take(8)
        .ToList();

    Console.WriteLine($"Most-recent 'Estate & Gift' projects ({eg.Count} shown):");
    foreach (var p in eg)
    {
        Console.WriteLine(
            $"    [{p.ProjectId}] {Truncate(p.ProjectName, 34)} | cust={Truncate(p.CustomerName, 24)} | " +
            $"PM={p.ProjectManager} | created={p.CreatedDateTime:o}");
    }

    Console.WriteLine();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Typed client failed: {ex.Message}");
    return 4;
}

// --- Step 4: team GI probe (property names + sample rows) -------------------
if (!string.IsNullOrWhiteSpace(options.TeamGenericInquiryName))
{
    try
    {
        var baseUrl = options.BaseUrl.TrimEnd('/');
        var teamUrl = $"{baseUrl}/t/{Uri.EscapeDataString(options.Tenant)}/api/odata/gi/" +
                      $"{Uri.EscapeDataString(options.TeamGenericInquiryName)}?$top=8";
        Console.WriteLine($"→ Probing team GI '{options.TeamGenericInquiryName}'…");
        using var req = new HttpRequestMessage(HttpMethod.Get, teamUrl);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        using var resp = await http.SendAsync(req, CancellationToken.None);
        var body = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode)
        {
            Console.Error.WriteLine($"❌ Team GI query failed ({(int)resp.StatusCode}): {Truncate(body, 300)}");
        }
        else
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var arr) && arr.GetArrayLength() > 0)
            {
                Console.WriteLine("✔ Team GI properties (name → sample value):");
                foreach (var prop in arr[0].EnumerateObject())
                {
                    var val = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() : prop.Value.GetRawText();
                    Console.WriteLine($"    {prop.Name,-28} → {Truncate(val, 50)}");
                }
                Console.WriteLine();
                Console.WriteLine("  Configured team field names:");
                foreach (var (label, field) in new[]
                {
                    ("TeamProjectIdField", options.TeamProjectIdField),
                    ("TeamEmailField", options.TeamEmailField),
                    ("TeamModifiedField", options.TeamModifiedField),
                })
                {
                    Console.WriteLine($"    {(arr[0].TryGetProperty(field, out _) ? "✔" : "✖")} {label,-20} = '{field}'");
                }
            }
            else
            {
                Console.WriteLine("⚠ Team GI returned no rows.");
            }
        }
        Console.WriteLine();
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"❌ Team GI probe failed: {ex.Message}");
    }
}

Console.WriteLine("✅ Acumatica connectivity OK.");
return 0;

// ---------------------------------------------------------------------------
static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!File.Exists(path))
    {
        return result;
    }

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

static bool Validate(AcumaticaOptions o, out string problem)
{
    if (string.IsNullOrWhiteSpace(o.BaseUrl) || o.BaseUrl.Contains("YOURCOMPANY"))
    {
        problem = "BaseUrl not set"; return false;
    }
    if (string.IsNullOrWhiteSpace(o.ClientId) || o.ClientId == "REPLACE_ME")
    {
        problem = "ClientId not set"; return false;
    }
    if (string.IsNullOrWhiteSpace(o.ClientSecret) || o.ClientSecret == "REPLACE_ME")
    {
        problem = "ClientSecret not set"; return false;
    }
    if (string.IsNullOrWhiteSpace(o.GenericInquiryName))
    {
        problem = "GenericInquiryName not set"; return false;
    }
    if (string.Equals(o.GrantType, "password", StringComparison.OrdinalIgnoreCase) &&
        (string.IsNullOrWhiteSpace(o.Username) || o.Username == "REPLACE_ME" ||
         string.IsNullOrWhiteSpace(o.Password) || o.Password == "REPLACE_ME"))
    {
        problem = "Username/Password not set (required for the password/ROPC grant)"; return false;
    }

    problem = string.Empty;
    return true;
}

static string Truncate(string? s, int max)
    => string.IsNullOrEmpty(s) ? string.Empty : (s.Length <= max ? s : s[..max] + "…");
