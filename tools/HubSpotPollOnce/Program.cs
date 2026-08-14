using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Acumatica;
using ProjectSync.HubSpot;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

// -----------------------------------------------------------------------------
// HubSpot scoping poll — DRY RUN plan.
//   Runs the watermark-based processor with a lookback window and prints the
//   scoping workspaces it WOULD create/update (deal id, resolved customer,
//   project name, practice, owner) — no SharePoint writes.
//   Usage: dotnet run --project tools/HubSpotPollOnce -- [lookbackHours]
// -----------------------------------------------------------------------------

var lookbackHours = args.Length > 0 && int.TryParse(args[0], out var lh) ? lh : 48;

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var options = new HubSpotOptions();
configuration.GetSection(HubSpotOptions.SectionName).Bind(options);
options.FirstRunLookbackHours = lookbackHours; // first run looks back this far so we see real deals

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; }));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(options.TimeoutSeconds) };
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

var wrapped = Options.Create(options);
var tokenProvider = new HubSpotTokenProvider(http, wrapped, loggerFactory.CreateLogger<HubSpotTokenProvider>());
var client = new HubSpotClient(http, tokenProvider, wrapped, loggerFactory.CreateLogger<HubSpotClient>());
var store = new InMemoryLastRunStore();
var processor = new HubSpotScopingProcessor(client, new NoOpSharePoint(), store, wrapped, TimeProvider.System, loggerFactory.CreateLogger<HubSpotScopingProcessor>());

Console.WriteLine($"=== HubSpot scoping poll (DRY RUN, lookback {lookbackHours}h) ===");
Console.WriteLine($"Practice scope : {(options.IncludedPractices.Count == 0 ? "<all>" : string.Join(", ", options.IncludedPractices))}");
Console.WriteLine($"Excluding stages: {string.Join(", ", options.TerminalStageIds)} (Won/Lost)");
Console.WriteLine();

var result = await processor.RunAsync(dryRun: true, CancellationToken.None);

Console.WriteLine($"Modified in window: {result.Found} | In-scope scoping deals: {result.InScope}");
Console.WriteLine();
Console.WriteLine("Scoping workspaces that WOULD be created/updated:");
Console.WriteLine($"  {"Deal Id",-13} {"Customer",-26} {"Project name",-34} {"Owner",-28}");
foreach (var p in result.Plan.OrderBy(p => p.CustomerName))
{
    Console.WriteLine($"  {p.DealId,-13} {Trunc(p.CustomerName, 26),-26} {Trunc(p.ProjectName, 34),-34} {p.OwnerEmail ?? "<none>",-28}");
}
Console.WriteLine();
Console.WriteLine($"✅ {result.Plan.Count} scoping workspace(s) planned. (Folder name would be first 10 of Customer + \" (dealId)\".)");
return 0;

static string Trunc(string? s, int max) => string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s[..(max - 1)] + "…");

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

// Dry-run only: the processor never calls SharePoint when dryRun=true.
file sealed class NoOpSharePoint : ISharePointDocumentSetService
{
    public Task<DocumentSetResult> EnsureProjectDocumentSetAsync(AcumaticaProject p, CancellationToken ct) => throw new NotSupportedException();
    public DocumentSetPlan PlanDocumentSet(AcumaticaProject p) => throw new NotSupportedException();
    public Task<DocumentSetResult> EnsureScopingWorkspaceAsync(ScopingWorkspace w, CancellationToken ct) => throw new NotSupportedException();
    public Task<ReconcileResult> ReconcileAsync(IReadOnlyList<AcumaticaProject> d, IReadOnlySet<string>? o, CancellationToken ct) => throw new NotSupportedException();
}

file sealed class InMemoryLastRunStore : ILastRunStore
{
    private readonly Dictionary<string, DateTimeOffset> _marks = new(StringComparer.OrdinalIgnoreCase);
    public Task<DateTimeOffset?> GetLastRunAsync(CancellationToken ct) => Task.FromResult<DateTimeOffset?>(null);
    public Task SetLastRunAsync(DateTimeOffset value, CancellationToken ct) => Task.CompletedTask;
    public Task<DateTimeOffset?> GetWatermarkAsync(string name, CancellationToken ct)
        => Task.FromResult(_marks.TryGetValue(name, out var v) ? v : (DateTimeOffset?)null);
    public Task SetWatermarkAsync(string name, DateTimeOffset value, CancellationToken ct)
    {
        _marks[name] = value;
        return Task.CompletedTask;
    }
}
