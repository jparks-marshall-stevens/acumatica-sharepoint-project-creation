using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Notifications;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Runs the client-upload scan once against the real SharePoint library and emails workspace members about
// any files uploaded in the lookback window. Gated by Notifications:TestMode, so it only reaches the test
// recipient. Use it to verify the feature against real data (e.g. after uploading a test file to a
// Client Uploads folder).
//
//   dotnet run --project tools/ClientUploadNotifyOnce -- [lookbackDays]   (default 7)
// -----------------------------------------------------------------------------

var lookbackDays = args.Select(a => a.Trim()).Where(a => int.TryParse(a, out _)).Select(int.Parse).FirstOrDefault(7);

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var notifyOptions = Bind<NotificationOptions>(configuration, NotificationOptions.SectionName);
var n = notifyOptions.Value;

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => o.SingleLine = true));

if (!n.TestMode || string.IsNullOrWhiteSpace(n.TestRecipient))
{
    Console.Error.WriteLine("❌ Refusing to run: Notifications:TestMode must be true and TestRecipient set.");
    return 1;
}

var since = DateTimeOffset.UtcNow.AddDays(-lookbackDays);
Console.WriteLine($"=== Client-upload scan (since {since:o}, ~{lookbackDays}d) — test send to {n.TestRecipient} ===");

var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var sender = new GraphEmailSender(contextFactory, spOptions, notifyOptions, loggerFactory.CreateLogger<GraphEmailSender>());
var notifier = new WorkspaceNotifier(sender, notifyOptions, loggerFactory.CreateLogger<WorkspaceNotifier>());
var service = new SharePointDocumentSetService(
    contextFactory,
    new GraphUploadLinkService(contextFactory, spOptions, loggerFactory.CreateLogger<GraphUploadLinkService>()),
    notifier, spOptions, loggerFactory.CreateLogger<SharePointDocumentSetService>());

var result = await service.ScanAndNotifyClientUploadsAsync(since, CancellationToken.None);

Console.WriteLine();
Console.WriteLine($"New files: {result.NewFiles}  |  Workspaces: {result.WorkspacesWithNewFiles}  |  Notified: {result.Notified}");
if (result.NewFiles == 0)
{
    Console.WriteLine("No client uploads found in the window. Upload a file to a 'Client Uploads' folder and re-run to test.");
}
return 0;

static IOptions<T> Bind<T>(IConfiguration config, string section) where T : class, new()
{
    var value = new T();
    config.GetSection(section).Bind(value);
    return Options.Create(value);
}

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!System.IO.File.Exists(path)) return result;
    using var doc = JsonDocument.Parse(System.IO.File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var values))
    {
        foreach (var prop in values.EnumerateObject())
        {
            result[prop.Name] = prop.Value.GetString();
        }
    }
    return result;
}
