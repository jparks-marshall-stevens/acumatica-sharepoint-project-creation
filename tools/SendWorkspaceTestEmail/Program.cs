using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Notifications;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Sends a notification email built from a REAL existing workspace, so the dataroom + client-upload links
// are the actual live URLs. Still gated by Notifications:TestMode — it only ever reaches TestRecipient.
//
//   dotnet run --project tools/SendWorkspaceTestEmail -- <match>
//     <match> = a value found in the OpportunityId, HubSpotDealId, or Project Id column, or a substring
//               of the folder name. Defaults to "PQ007180" (Blackstone Dilworth).
// -----------------------------------------------------------------------------

var match = args.FirstOrDefault(a => !a.StartsWith("--"))?.Trim() ?? "PQ007180";

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var notifyOptions = Bind<NotificationOptions>(configuration, NotificationOptions.SectionName);
var sp = spOptions.Value;
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

var mapping = sp.PracticeMappings.First();
var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? sp.SiteUrl : mapping.SiteUrl!;
var origin = new Uri(siteUrl).GetLeftPart(UriPartial.Authority);

var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine($"=== Send REAL-workspace test email (match '{match}') ===");
Console.WriteLine($"Site: {siteUrl}  |  Library: {mapping.Library}  |  To (test): {n.TestRecipient}");
Console.WriteLine();

using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(mapping.Library);
ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
await ctx.ExecuteQueryRetryAsync();

// Scan document sets for one matching the argument in any key column or the folder name.
// Match on METADATA columns only (never the file path) so we hit the document set, not a subfolder
// like "Client Uploads" whose path happens to contain the search term.
string?[] cols =
{
    sp.OpportunityIdColumn, sp.HubSpotDealIdColumn, sp.ProjectIdColumn, sp.CustomerNameColumn,
    sp.ProjectNameColumn,
};

ListItem? hit = null;
ListItemCollectionPosition? position = null;
do
{
    var query = new CamlQuery
    {
        ViewXml =
            "<View Scope='RecursiveAll'><Query><Where>" +
            "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
            "</Where></Query><RowLimit Paged='TRUE'>500</RowLimit></View>",
        ListItemCollectionPosition = position,
    };
    var items = list.GetItems(query);
    ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(
        i => i["FileRef"], i => i["FileLeafRef"], i => i[sp.CustomerNameColumn], i => i[sp.ProjectNameColumn],
        i => i[sp.ProjectIdColumn], i => i[sp.OpportunityIdColumn], i => i[sp.HubSpotDealIdColumn],
        i => i[sp.StatusColumn], i => i[sp.ClientUploadLinkColumn]));
    await ctx.ExecuteQueryRetryAsync();

    foreach (var it in items)
    {
        bool Matches(string? col) =>
            !string.IsNullOrWhiteSpace(col) &&
            it.FieldValues.TryGetValue(col!, out var v) &&
            (v?.ToString() ?? string.Empty).Contains(match, StringComparison.OrdinalIgnoreCase);

        if (cols.Any(Matches))
        {
            hit = it;
            break;
        }
    }

    position = hit is null ? items.ListItemCollectionPosition : null;
}
while (position is not null);

if (hit is null)
{
    Console.Error.WriteLine($"❌ No document set matched '{match}'.");
    return 2;
}

string? F(string? col) => !string.IsNullOrWhiteSpace(col) && hit.FieldValues.TryGetValue(col!, out var v) ? v?.ToString() : null;

var fileRef = F("FileRef") ?? string.Empty;
var status = F(sp.StatusColumn);
var isScoping = string.Equals(status, sp.ScopingStatusValue, StringComparison.OrdinalIgnoreCase);
var dataroomUrl = origin + fileRef.Replace(" ", "%20");
var uploadLink = F(sp.ClientUploadLinkColumn);

var notice = new WorkspaceNotice
{
    Phase = isScoping ? WorkspacePhase.Scoping : WorkspacePhase.Execution,
    CustomerName = F(sp.CustomerNameColumn) ?? "(unknown)",
    EngagementName = F(sp.ProjectNameColumn),
    IdLabel = isScoping ? "Opportunity #" : "Project ID",
    IdValue = isScoping ? F(sp.OpportunityIdColumn) : F(sp.ProjectIdColumn),
    ProjectManager = null,
    Practice = mapping.Practice,
    DataroomUrl = dataroomUrl,
    UploadLinkUrl = uploadLink,
};

Console.WriteLine("Matched workspace — REAL values that will be in the email:");
Console.WriteLine($"    Folder      : {F("FileLeafRef")}");
Console.WriteLine($"    Status      : {status}  (phase: {notice.Phase})");
Console.WriteLine($"    Customer    : {notice.CustomerName}");
Console.WriteLine($"    Dataroom URL: {dataroomUrl}");
Console.WriteLine($"    Upload link : {(string.IsNullOrWhiteSpace(uploadLink) ? "<none stored>" : uploadLink)}");
Console.WriteLine();

var sender = new GraphEmailSender(contextFactory, spOptions, notifyOptions, loggerFactory.CreateLogger<GraphEmailSender>());
var notifier = new WorkspaceNotifier(sender, notifyOptions, loggerFactory.CreateLogger<WorkspaceNotifier>());

// --upload: send the CLIENT-UPLOAD email with the real uploads-folder URL (and the folder's real files).
if (args.Any(a => a.Equals("--upload", StringComparison.OrdinalIgnoreCase)))
{
    var uploadsRel = fileRef + "/" + sp.ClientUploadsFolderName;
    var uploadsFolderUrl = origin + uploadsRel.Replace(" ", "%20");

    var names = new List<string>();
    try
    {
        var uf = ctx.Web.GetFolderByServerRelativeUrl(uploadsRel);
        ctx.Load(uf.Files, fs => fs.Include(f => f.Name));
        await ctx.ExecuteQueryRetryAsync();
        names = uf.Files.Select(f => f.Name).ToList();
    }
    catch { /* folder may not exist yet */ }

    if (names.Count == 0)
    {
        names = new List<string> { "Trust Agreement.pdf", "2024 Financial Statements.xlsx" };
        Console.WriteLine("(Client Uploads folder is empty — using sample filenames; the button URL is still real.)");
    }

    var uploadRecipients = new List<string?> { "someone.live@example.com", mapping.PracticeLeaderEmail };
    uploadRecipients.AddRange(mapping.AdminEmails);
    var exclude = isScoping ? null : mapping.PracticeLeaderEmail; // keep Bruce for scoping, drop for engagement

    Console.WriteLine($"Uploads folder URL: {uploadsFolderUrl}");
    Console.WriteLine($"Files: {string.Join(", ", names)}");
    await notifier.NotifyClientUploadAsync(notice, names, uploadsFolderUrl, uploadRecipients, exclude, CancellationToken.None);
    Console.WriteLine($"✔ Sent client-upload test to {n.TestRecipient}. Click 'Open the client uploads' to verify it resolves.");
    return 0;
}

// Recipients mirror production: a sample owner/PM (stand-in for the real one) PLUS the practice admins
// from config (e.g. Michelle). TestMode still routes everything to TestRecipient only; the intended list
// in the subject shows who would really receive it.
var recipients = new List<string?>();
recipients.Add(isScoping ? "deal.owner.sample@example.com" : "project.pm.sample@example.com");
recipients.AddRange(mapping.AdminEmails);
await notifier.NotifyCreatedAsync(notice, recipients, mapping.PracticeLeaderEmail, CancellationToken.None);

Console.WriteLine();
Console.WriteLine($"✔ Sent to {n.TestRecipient}. Click 'Open the dataroom' and 'Client file-request link' to verify both resolve.");
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
