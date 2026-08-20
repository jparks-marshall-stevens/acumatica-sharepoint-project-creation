using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// One-time: renames existing document-set folders to the current naming convention
// (BuildDocumentSetName with the configured DocumentSetNameMaxLength), then re-syncs the changed dataroom
// URL to Acumatica (DATAURL) and the DataroomUrl metadata column. Client-upload links + permissions + item
// id survive a rename, so only the path/URL changes.
//
//   dotnet run --project tools/SyncFolderNames -- --dry     (preview renames, no changes)
//   dotnet run --project tools/SyncFolderNames --            (execute)
// -----------------------------------------------------------------------------

var dry = args.Any(a => a.Equals("--dry", StringComparison.OrdinalIgnoreCase));

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));
var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath)).AddEnvironmentVariables().Build();

var acuOptions = Bind<AcumaticaOptions>(configuration, AcumaticaOptions.SectionName);
var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var sp = spOptions.Value;

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Warning)
    .AddFilter("ProjectSync", Microsoft.Extensions.Logging.LogLevel.Warning)
    .AddSimpleConsole(o => o.SingleLine = true));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acuOptions.Value.TimeoutSeconds) };
var tokenProvider = new AcumaticaTokenProvider(http, acuOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acuOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());

var mapping = sp.PracticeMappings.First();
var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? sp.SiteUrl : mapping.SiteUrl!;
var origin = new Uri(siteUrl).GetLeftPart(UriPartial.Authority);
var pidCol = sp.ProjectIdColumn;
var custCol = sp.CustomerNameColumn;
var urlCol = sp.DataroomUrlColumn;

Console.WriteLine($"Mode        : {(dry ? "DRY RUN (no changes)" : "EXECUTE")}");
Console.WriteLine($"Library     : {mapping.Library}   |   name max length: {sp.DocumentSetNameMaxLength}");
Console.WriteLine();

using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(mapping.Library);
ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
await ctx.ExecuteQueryRetryAsync();

var rows = new List<(string ProjectId, string FileRef, string? Customer)>();
ListItemCollectionPosition? position = null;
do
{
    var query = new CamlQuery
    {
        ViewXml =
            "<View Scope='RecursiveAll'><Query><Where><And>" +
            "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
            $"<IsNotNull><FieldRef Name='{pidCol}'/></IsNotNull>" +
            "</And></Where></Query>" +
            $"<ViewFields><FieldRef Name='{pidCol}'/><FieldRef Name='FileRef'/><FieldRef Name='FileLeafRef'/><FieldRef Name='{custCol}'/></ViewFields>" +
            "<RowLimit Paged='TRUE'>2000</RowLimit></View>",
        ListItemCollectionPosition = position,
    };
    var items = list.GetItems(query);
    ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(i => i["FileRef"], i => i[pidCol], i => i[custCol]));
    await ctx.ExecuteQueryRetryAsync();
    foreach (var it in items)
    {
        var pid = it[pidCol]?.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(pid)) continue;
        rows.Add((pid!, it["FileRef"]?.ToString() ?? "", it.FieldValues.TryGetValue(custCol, out var cv) ? cv?.ToString() : null));
    }
    position = items.ListItemCollectionPosition;
}
while (position is not null);

Console.WriteLine($"Found {rows.Count} document set(s) with a Project Id.");
Console.WriteLine();

int renamed = 0, resynced = 0, resyncFailed = 0, unchanged = 0;
foreach (var (pid, fileRef, customer) in rows.OrderBy(r => r.ProjectId, StringComparer.OrdinalIgnoreCase))
{
    var currentLeaf = fileRef.TrimEnd('/').Split('/').Last();
    var desired = SharePointNaming.BuildDocumentSetName(customer, pid, sp.DocumentSetNameMaxLength);

    if (string.Equals(currentLeaf, desired, StringComparison.Ordinal))
    {
        unchanged++;
        continue;
    }

    Console.WriteLine($"  {pid}");
    Console.WriteLine($"      from: {currentLeaf}");
    Console.WriteLine($"      to  : {desired}");
    if (dry) continue;

    try
    {
        // Rename via FileLeafRef (a rename, not a move — item id, permissions, and sharing links follow).
        var item = ctx.Web.GetFolderByServerRelativeUrl(fileRef).ListItemAllFields;
        item["FileLeafRef"] = desired;
        item.Update();
        await ctx.ExecuteQueryRetryAsync();
        ctx.Load(item, i => i["FileRef"]);
        await ctx.ExecuteQueryRetryAsync();
        renamed++;

        var newFileRef = item["FileRef"]?.ToString() ?? fileRef;
        var newDataUrl = origin + newFileRef.Replace(" ", "%20");

        // Re-sync the dataroom URL to Acumatica (DATAURL) and the metadata column.
        bool ok = false;
        try { ok = await acumatica.WriteProjectUrlsAsync(pid, newDataUrl, clientUrl: null, CancellationToken.None); }
        catch (Exception ex) { Console.WriteLine($"      write error: {ex.Message}"); }
        if (ok) resynced++; else resyncFailed++;

        try
        {
            var it2 = ctx.Web.GetFolderByServerRelativeUrl(newFileRef).ListItemAllFields;
            ctx.Load(it2);
            await ctx.ExecuteQueryRetryAsync();
            if (it2.FieldValues.ContainsKey(urlCol)) { it2[urlCol] = newDataUrl; it2.Update(); await ctx.ExecuteQueryRetryAsync(); }
        }
        catch (Exception ex) { Console.WriteLine($"      stamp warning: {ex.Message}"); }

        Console.WriteLine($"      done (DATAURL {(ok ? "updated" : "NOT updated")})");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"      RENAME FAILED: {ex.Message}");
    }
}

Console.WriteLine();
Console.WriteLine(dry
    ? $"DRY RUN — {rows.Count - unchanged} folder(s) would be renamed, {unchanged} already correct. Re-run without --dry to execute."
    : $"Done — renamed {renamed}, DATAURL re-synced {resynced} (failed {resyncFailed}), already-correct {unchanged}.");
return resyncFailed > 0 ? 3 : 0;

static IOptions<T> Bind<T>(IConfiguration c, string s) where T : class, new()
{ var v = new T(); c.GetSection(s).Bind(v); return Options.Create(v); }
static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var r = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!System.IO.File.Exists(path)) return r;
    using var doc = JsonDocument.Parse(System.IO.File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var vals))
        foreach (var p in vals.EnumerateObject()) r[p.Name] = p.Value.GetString();
    return r;
}
