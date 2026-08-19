using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Notifications;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Sends one of each notification email through the REAL Graph sender, so we can confirm the whole path
// (cert auth -> Mail.Send -> Exchange Application Access Policy -> rendering) end to end.
//
// It deliberately passes fake "live" recipients to prove TestMode redirects everything to the single
// TestRecipient. With Notifications:TestMode=true (the default in local.settings), NOTHING reaches a live
// mailbox — every message goes only to Notifications:TestRecipient.
//
//   dotnet run --project tools/SendTestEmail
// -----------------------------------------------------------------------------

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

Console.WriteLine("=== Send test notification emails ===");
Console.WriteLine($"From        : {(string.IsNullOrWhiteSpace(n.FromAddress) ? "<NOT SET>" : n.FromAddress)}");
Console.WriteLine($"TestMode    : {n.TestMode}");
Console.WriteLine($"TestRecipient: {n.TestRecipient}");
Console.WriteLine($"LogoUrl     : {(string.IsNullOrWhiteSpace(n.LogoUrl) ? "<none>" : n.LogoUrl)}");
Console.WriteLine();

if (!n.TestMode)
{
    Console.Error.WriteLine("❌ Refusing to run: Notifications:TestMode is not true. This tool is for safe test sends only.");
    return 1;
}
if (string.IsNullOrWhiteSpace(n.TestRecipient))
{
    Console.Error.WriteLine("❌ Notifications:TestRecipient is blank.");
    return 1;
}

var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var sender = new GraphEmailSender(contextFactory, spOptions, notifyOptions, loggerFactory.CreateLogger<GraphEmailSender>());
var notifier = new WorkspaceNotifier(sender, notifyOptions, loggerFactory.CreateLogger<WorkspaceNotifier>());

// Fake "live" recipients on purpose — TestMode must redirect all of these to the TestRecipient.
var fakeLive = new[] { "someone.live@example.com", "another.live@example.com" };
const string leader = "bjohnson@marshall-stevens.com";

var scoping = new WorkspaceNotice
{
    Phase = WorkspacePhase.Scoping,
    CustomerName = "Blackstone Dilworth",
    EngagementName = "Estate valuation — Dilworth family trust",
    IdLabel = "Opportunity #",
    IdValue = "PQ007180",
    Practice = "Estate & Gift",
    DataroomUrl = "https://marshallstevens.sharepoint.com/sites/GiftEstate/Shared%20Documents/Projects/Current/Blackstone%20Dilworth%20(PQ007180)",
    UploadLinkUrl = "https://marshallstevens.sharepoint.com/:f:/s/GiftEstate/EXAMPLE-upload-link",
};

var execution = new WorkspaceNotice
{
    Phase = WorkspacePhase.Execution,
    CustomerName = "Robert Palmer",
    EngagementName = "Gift valuation of a minority interest",
    IdLabel = "Project ID",
    IdValue = "10-31-21-74663",
    ProjectManager = "Matthew West",
    Practice = "Estate & Gift",
    DataroomUrl = "https://marshallstevens.sharepoint.com/sites/GiftEstate/Shared%20Documents/Projects/Current/Robert%20Palmer%20(10-31-21-74663)",
    UploadLinkUrl = "https://marshallstevens.sharepoint.com/:f:/s/GiftEstate/EXAMPLE-upload-link",
};

Console.WriteLine("Sending: scoping-created…");
await notifier.NotifyCreatedAsync(scoping, fakeLive, leader, CancellationToken.None);

Console.WriteLine("Sending: project-created…");
await notifier.NotifyCreatedAsync(execution, fakeLive, leader, CancellationToken.None);

Console.WriteLine("Sending: access-added…");
await notifier.NotifyAccessAddedAsync(execution, new[] { "someone.live@example.com" }, leader, CancellationToken.None);

Console.WriteLine();
Console.WriteLine($"✔ Done. If successful, three '[TEST]' emails are in {n.TestRecipient}'s inbox.");
Console.WriteLine("  (Look for a 'Sent ... as ...' line above for each; a 'Graph sendMail failed' line means check the error.)");
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
