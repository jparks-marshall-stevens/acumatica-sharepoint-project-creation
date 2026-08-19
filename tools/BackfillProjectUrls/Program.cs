using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// One-time backfill: for every document set in the Gift & Estate library that has a Project Id, write the
// CURRENT dataroom URL + client-upload URL back to Acumatica (DATAURL / CLIENTURL), correcting any stale
// values, and stamp the DataroomUrl metadata column so the reconcile baseline matches (no later re-sync).
//
//   dotnet run --project tools/BackfillProjectUrls -- --dry     (list only, no writes)
//   dotnet run --project tools/BackfillProjectUrls --            (execute)
// -----------------------------------------------------------------------------

var dry = args.Any(a => a.Equals("--dry", StringComparison.OrdinalIgnoreCase));

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));
var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables().Build();

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
var linkCol = sp.ClientUploadLinkColumn;
var urlCol = sp.DataroomUrlColumn;

Console.WriteLine($"Mode        : {(dry ? "DRY RUN (no writes)" : "EXECUTE")}");
Console.WriteLine($"Site        : {siteUrl}");
Console.WriteLine($"Library     : {mapping.Library}");
Console.WriteLine();

using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(mapping.Library);
ctx.Load(list, l => l.Fields.Include(f => f.InternalName));
await ctx.ExecuteQueryRetryAsync();

// Ensure the DataroomUrl column exists so we can stamp it.
if (!dry && !list.Fields.Any(f => f.InternalName == urlCol))
{
    list.Fields.AddFieldAsXml(
        $"<Field Type='Text' Name='{urlCol}' StaticName='{urlCol}' DisplayName='{urlCol}' Group='ProjectSync'/>",
        addToDefaultView: false, options: AddFieldOptions.AddFieldInternalNameHint);
    await ctx.ExecuteQueryRetryAsync();
    Console.WriteLine($"Created metadata column '{urlCol}'.");
}
var hasUrlCol = dry ? list.Fields.Any(f => f.InternalName == urlCol) : true;

// Page through every document-set folder that carries a Project Id.
var rows = new List<(string ProjectId, string FileRef, string? ClientUrl, string? StoredUrl)>();
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
            $"<ViewFields><FieldRef Name='{pidCol}'/><FieldRef Name='FileRef'/><FieldRef Name='{linkCol}'/>" +
            (hasUrlCol ? $"<FieldRef Name='{urlCol}'/>" : "") +
            "</ViewFields><RowLimit Paged='TRUE'>2000</RowLimit></View>",
        ListItemCollectionPosition = position,
    };
    var items = list.GetItems(query);
    ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(i => i["FileRef"], i => i[pidCol], i => i[linkCol]));
    if (hasUrlCol) ctx.Load(items, c => c.Include(i => i[urlCol]));
    await ctx.ExecuteQueryRetryAsync();

    foreach (var it in items)
    {
        var pid = it[pidCol]?.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(pid)) continue;
        var fileRef = it["FileRef"]?.ToString() ?? "";
        var clientUrl = it.FieldValues.TryGetValue(linkCol, out var cv) ? cv?.ToString() : null;
        var storedUrl = hasUrlCol && it.FieldValues.TryGetValue(urlCol, out var sv) ? sv?.ToString() : null;
        rows.Add((pid!, fileRef, clientUrl, storedUrl));
    }
    position = items.ListItemCollectionPosition;
}
while (position is not null);

Console.WriteLine($"Found {rows.Count} document set(s) with a Project Id.");
Console.WriteLine();

int written = 0, failed = 0, skippedDisabled = 0;
var writeBackDisabled = string.IsNullOrWhiteSpace(acuOptions.Value.DataUrlAttributeId) && string.IsNullOrWhiteSpace(acuOptions.Value.ClientUrlAttributeId);
if (writeBackDisabled)
{
    Console.WriteLine("⚠ Write-back is disabled (no DataUrlAttributeId/ClientUrlAttributeId configured). Nothing will be written to Acumatica.");
}

foreach (var (pid, fileRef, clientUrl, storedUrl) in rows.OrderBy(r => r.ProjectId, StringComparer.OrdinalIgnoreCase))
{
    var dataUrl = origin + fileRef.Replace(" ", "%20");
    var currentlyStale = !string.IsNullOrWhiteSpace(storedUrl) && !string.Equals(storedUrl, dataUrl, StringComparison.OrdinalIgnoreCase);
    var flag = string.IsNullOrWhiteSpace(storedUrl) ? "new " : currentlyStale ? "FIX " : "ok  ";
    Console.WriteLine($"  [{flag}] {pid}  ->  {dataUrl}{(string.IsNullOrWhiteSpace(clientUrl) ? "  (no client link)" : "")}");

    if (dry) continue;

    // 1) write to Acumatica (DATAURL + CLIENTURL)
    bool ok = false;
    if (!writeBackDisabled)
    {
        try { ok = await acumatica.WriteProjectUrlsAsync(pid, dataUrl, clientUrl, CancellationToken.None); }
        catch (Exception ex) { Console.WriteLine($"       write error: {ex.Message}"); }
        if (ok) written++; else { failed++; Console.WriteLine("       ✘ Acumatica write returned false/failed"); }
    }
    else { skippedDisabled++; }

    // 2) stamp the DataroomUrl metadata column so reconcile won't re-sync it
    try
    {
        var item = ctx.Web.GetFolderByServerRelativeUrl(fileRef).ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();
        if (item.FieldValues.ContainsKey(urlCol)) { item[urlCol] = dataUrl; item.Update(); await ctx.ExecuteQueryRetryAsync(); }
    }
    catch (Exception ex) { Console.WriteLine($"       stamp warning: {ex.Message}"); }
}

Console.WriteLine();
Console.WriteLine(dry
    ? $"DRY RUN complete — {rows.Count} project(s) would be processed. Re-run without --dry to execute."
    : $"Backfill complete — Acumatica writes ok: {written}, failed: {failed}{(skippedDisabled > 0 ? $", skipped(disabled): {skippedDisabled}" : "")}.");
return failed > 0 ? 3 : 0;

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
