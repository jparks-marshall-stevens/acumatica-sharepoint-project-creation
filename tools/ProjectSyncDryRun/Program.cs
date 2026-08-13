using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

// -----------------------------------------------------------------------------
// ProjectSync DRY RUN.
//
// Runs the real ProjectSyncProcessor in dry-run mode: pulls projects from the
// Acumatica GI, applies the exclusion + practice filters, and reports exactly
// which SharePoint document sets WOULD be created and where — creating nothing
// and connecting to SharePoint not at all.
//
// Needs no Azure Functions runtime / Azurite: uses an in-memory last-run store.
// Config is read from the Functions local.settings.json "Values" section.
//
// Usage: dotnet run --project tools/ProjectSyncDryRun [days]
//        (days = look-back window; default 30)
// -----------------------------------------------------------------------------

var days = args.Length > 0 && int.TryParse(args[0], out var d) ? d : 30;

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var acumaticaOptions = Bind<AcumaticaOptions>(configuration, AcumaticaOptions.SectionName);
var sharePointOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var stateOptions = Bind<StateOptions>(configuration, StateOptions.SectionName);

using var loggerFactory = LoggerFactory.Create(b => b.SetMinimumLevel(LogLevel.Warning).AddSimpleConsole(o => o.SingleLine = true));
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acumaticaOptions.Value.TimeoutSeconds) };

var tokenProvider = new AcumaticaTokenProvider(http, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(sharePointOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var sharePoint = new SharePointDocumentSetService(contextFactory, sharePointOptions, loggerFactory.CreateLogger<SharePointDocumentSetService>());

var processor = new ProjectSyncProcessor(
    acumatica, sharePoint, new InMemoryLastRunStore(), stateOptions, acumaticaOptions,
    TimeProvider.System, loggerFactory.CreateLogger<ProjectSyncProcessor>());

Console.WriteLine("=== ProjectSync DRY RUN ===");
Console.WriteLine($"GI                : {acumaticaOptions.Value.GenericInquiryName}");
Console.WriteLine($"Included practices: {string.Join(", ", acumaticaOptions.Value.IncludedPractices)}");
Console.WriteLine($"Excluded project ids: {string.Join(", ", acumaticaOptions.Value.ExcludedProjectIds)}");
Console.WriteLine($"Window            : projects created in the last {days} day(s)");
Console.WriteLine();

var result = await processor.RunAsync(
    new RunOptions { DryRun = true, OverrideSince = DateTimeOffset.UtcNow.AddDays(-days) },
    CancellationToken.None);

Console.WriteLine($"GI returned : {result.Found}");
Console.WriteLine($"Skipped     : {result.Skipped}  (excluded id or practice not in allow-list)");
Console.WriteLine($"WOULD CREATE: {result.Planned}");
Console.WriteLine();

if (result.Plan is { Count: > 0 })
{
    var rows = result.Plan.OrderByDescending(x => x.CreatedDateTime).ToList();

    Console.WriteLine($"All sets land in: {rows[0].TargetLibrary}/{rows[0].TargetFolder}  @ {rows[0].TargetSiteUrl}");
    Console.WriteLine();

    // Aligned table of the values that would be written.
    const int wName = 42, wId = 16, wCust = 30, wPm = 20, wDate = 10;
    string H(string s, int w) => (s.Length > w ? s[..(w - 1)] + "…" : s).PadRight(w);
    Console.WriteLine($"{H("Folder name (= first 40 of Description)", wName)} {H("Project Id", wId)} {H("Customer", wCust)} {H("Project Manager", wPm)} {H("Created", wDate)}");
    Console.WriteLine(new string('-', wName + wId + wCust + wPm + wDate + 4));
    foreach (var p in rows)
    {
        Console.WriteLine(
            $"{H(p.DocumentSetName ?? "", wName)} {H(p.ProjectId, wId)} {H((p.CustomerName ?? "").Trim(), wCust)} " +
            $"{H(p.ProjectManager ?? "", wPm)} {H(p.CreatedDateTime?.ToString("yyyy-MM-dd") ?? "", wDate)}");
    }
    Console.WriteLine();
    Console.WriteLine("Note: identical folder names (same 40-char prefix) get the Project Id appended at creation time.");

    // People-field resolution outlook: PM emails outside your M365 tenant won't resolve.
    const string tenantDomain = "marshall-stevens.com";
    Console.WriteLine();
    Console.WriteLine("Project Manager email domains (People field resolves only in-tenant):");
    foreach (var g in rows
        .Select(r => (r.ProjectManagerEmail ?? r.ProjectManager ?? "").Split('@').LastOrDefault()?.ToLowerInvariant() ?? "")
        .GroupBy(dom => string.IsNullOrWhiteSpace(dom) ? "<blank>" : dom)
        .OrderByDescending(g => g.Count()))
    {
        var mark = g.Key == tenantDomain ? "✔ resolves" : "✖ left blank";
        Console.WriteLine($"    {g.Key,-28} → {g.Count(),3}   {mark}");
    }

    // Detail: the specific projects whose PM is NOT on the tenant domain (candidates for mapping).
    string DomainOf(PlannedDocumentSet r) => (r.ProjectManagerEmail ?? r.ProjectManager ?? "").Split('@').LastOrDefault()?.ToLowerInvariant() ?? "";
    var external = rows.Where(r => DomainOf(r) != tenantDomain).ToList();
    if (external.Count > 0)
    {
        Console.WriteLine();
        Console.WriteLine($"Projects with a non-{tenantDomain} PM ({external.Count}):");
        foreach (var r in external)
        {
            Console.WriteLine($"    {r.CreatedDateTime:yyyy-MM-dd}  {r.ProjectId,-16} {(r.ProjectManagerEmail ?? r.ProjectManager),-34} {r.CustomerName?.Trim()}");
        }
    }
}
else
{
    Console.WriteLine("Nothing would be created in this window.");
}

Console.WriteLine();
Console.WriteLine("(dry run — nothing was created, watermark unchanged)");
return 0;

// ---------------------------------------------------------------------------
static IOptions<T> Bind<T>(IConfiguration config, string section) where T : class, new()
{
    var value = new T();
    config.GetSection(section).Bind(value);
    return Microsoft.Extensions.Options.Options.Create(value);
}

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

/// <summary>No-op last-run store: dry-run reads null (uses the OverrideSince window) and never writes.</summary>
file sealed class InMemoryLastRunStore : ILastRunStore
{
    public Task<DateTimeOffset?> GetLastRunAsync(CancellationToken cancellationToken) => Task.FromResult<DateTimeOffset?>(null);
    public Task SetLastRunAsync(DateTimeOffset value, CancellationToken cancellationToken) => Task.CompletedTask;
}
