using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

// Runs one reconcile pass against the real systems. "full" (default) sweeps all tracked sets;
// "incremental" uses the team GI's modified date. Reads config from local.settings.json.
// Usage: dotnet run --project tools/ReconcileOnce [full|incremental]

var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "full";

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

using var loggerFactory = LoggerFactory.Create(b => b.SetMinimumLevel(LogLevel.Information).AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; }));
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acumaticaOptions.Value.TimeoutSeconds) };

var tokenProvider = new AcumaticaTokenProvider(http, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(sharePointOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var uploadLinks = new GraphUploadLinkService(contextFactory, sharePointOptions, loggerFactory.CreateLogger<GraphUploadLinkService>());
var sharePoint = new SharePointDocumentSetService(contextFactory, uploadLinks, sharePointOptions, loggerFactory.CreateLogger<SharePointDocumentSetService>());
var processor = new ProjectSyncProcessor(
    acumatica, sharePoint, new InMemoryLastRunStore(), stateOptions, acumaticaOptions,
    TimeProvider.System, loggerFactory.CreateLogger<ProjectSyncProcessor>());

Console.WriteLine($"=== Reconcile ({mode}) ===");
var result = mode == "incremental"
    ? await processor.ReconcileIncrementalAsync(CancellationToken.None)
    : await processor.ReconcileFullAsync(CancellationToken.None);

Console.WriteLine();
Console.WriteLine($"Considered : {result.Considered}");
Console.WriteLine($"Updated    : {result.Updated}");
Console.WriteLine($"Unchanged  : {result.Unchanged}");
Console.WriteLine($"NotTracked : {result.NotTracked}");
return 0;

static IOptions<T> Bind<T>(IConfiguration config, string section) where T : class, new()
{
    var value = new T();
    config.GetSection(section).Bind(value);
    return Microsoft.Extensions.Options.Options.Create(value);
}

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!File.Exists(path)) return result;
    using var doc = JsonDocument.Parse(File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var values))
        foreach (var prop in values.EnumerateObject())
            result[prop.Name] = prop.Value.GetString();
    return result;
}

file sealed class InMemoryLastRunStore : ILastRunStore
{
    public Task<DateTimeOffset?> GetLastRunAsync(CancellationToken ct) => Task.FromResult<DateTimeOffset?>(null);
    public Task SetLastRunAsync(DateTimeOffset value, CancellationToken ct) => Task.CompletedTask;
    public Task<DateTimeOffset?> GetWatermarkAsync(string name, CancellationToken ct) => Task.FromResult<DateTimeOffset?>(null);
    public Task SetWatermarkAsync(string name, DateTimeOffset value, CancellationToken ct) => Task.CompletedTask;
}
