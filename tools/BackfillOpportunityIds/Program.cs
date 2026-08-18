using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.HubSpot;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// -----------------------------------------------------------------------------
// Backfill OpportunityId (and folder names) on existing scoping workspaces.
//
// Workspaces created before the opportunity number was wired up carry a HubSpotDealId but a blank
// OpportunityId — so an Acumatica PQCode has nothing to match and the workspace would be duplicated
// instead of promoted. This re-reads each deal's opportunity number from HubSpot (batch read by id, so
// deals that have since closed are still found), stamps the column, and renames folders still named after
// the raw deal record id to the "{customer} ({PQCode})" form.
//
// Touches the OpportunityId column and the folder leaf name only — never Status, permissions, or other
// metadata. Renames are in-place (FileLeafRef): contents, item id, and sharing links follow the item.
//
//   dotnet run --project tools/BackfillOpportunityIds              → dry run (default; writes nothing)
//   dotnet run --project tools/BackfillOpportunityIds -- --apply   → write the column
// -----------------------------------------------------------------------------

var apply = args.Contains("--apply", StringComparer.OrdinalIgnoreCase);

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var hubOptions = Bind<HubSpotOptions>(configuration, HubSpotOptions.SectionName);
var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var sp = spOptions.Value;

if (string.IsNullOrWhiteSpace(sp.OpportunityIdColumn) || string.IsNullOrWhiteSpace(sp.HubSpotDealIdColumn))
{
    Console.Error.WriteLine("❌ SharePoint:OpportunityIdColumn and SharePoint:HubSpotDealIdColumn must both be set.");
    return 1;
}

if (string.IsNullOrWhiteSpace(hubOptions.Value.OpportunityIdProperty))
{
    Console.Error.WriteLine("❌ HubSpot:OpportunityIdProperty is blank — nothing to read the opportunity number from.");
    return 1;
}

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => o.SingleLine = true));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(hubOptions.Value.TimeoutSeconds) };
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

var tokenProvider = new HubSpotTokenProvider(http, hubOptions, loggerFactory.CreateLogger<HubSpotTokenProvider>());
var hubspot = new HubSpotClient(http, tokenProvider, hubOptions, loggerFactory.CreateLogger<HubSpotClient>());
var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine($"=== Backfill {sp.OpportunityIdColumn} on scoping workspaces{(apply ? "" : "  (DRY RUN)")} ===");
Console.WriteLine();

var mapping = sp.PracticeMappings.FirstOrDefault();
if (mapping is null)
{
    Console.Error.WriteLine("❌ No SharePoint:PracticeMappings configured.");
    return 1;
}

var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? sp.SiteUrl : mapping.SiteUrl!;
Console.WriteLine($"Site    : {siteUrl}");
Console.WriteLine($"Library : {mapping.Library}");
Console.WriteLine();

using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(mapping.Library);
ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName));
await ctx.ExecuteQueryRetryAsync();

// The sync auto-creates this column, but it may not have run yet against this library. Reading a
// non-existent column throws, so establish whether it is there before building the query.
var hasOpportunityColumn = list.Fields.Any(f => f.InternalName == sp.OpportunityIdColumn);
if (!hasOpportunityColumn)
{
    if (!apply)
    {
        Console.WriteLine($"⚠ Column '{sp.OpportunityIdColumn}' does not exist yet — it will be created on --apply.");
        Console.WriteLine("  Every workspace below therefore counts as blank.");
        Console.WriteLine();
    }
    else
    {
        Console.WriteLine($"Creating column '{sp.OpportunityIdColumn}'…");
        list.Fields.AddFieldAsXml(
            $"<Field Type='Text' Name='{sp.OpportunityIdColumn}' StaticName='{sp.OpportunityIdColumn}' " +
            $"DisplayName='{sp.OpportunityIdColumn}' Group='ProjectSync'/>",
            addToDefaultView: true, options: AddFieldOptions.AddFieldInternalNameHint);
        await ctx.ExecuteQueryRetryAsync();
        hasOpportunityColumn = true;
        Console.WriteLine("✔ Created.");
        Console.WriteLine();
    }
}

// Every document set carrying a HubSpot deal id, paged. Includes already-promoted ones: stamping the
// opportunity number on those is harmless and keeps the deal↔project link legible.
var sets = new List<(string Url, string Leaf, string DealId, string? Opportunity,
                    string? Customer, string? ProjectId, string? Status)>();
ListItemCollectionPosition? position = null;
do
{
    var query = new CamlQuery
    {
        ViewXml =
            "<View Scope='RecursiveAll'><Query><Where><And>" +
            "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
            $"<IsNotNull><FieldRef Name='{sp.HubSpotDealIdColumn}'/></IsNotNull>" +
            "</And></Where></Query>" +
            $"<ViewFields><FieldRef Name='FileRef'/><FieldRef Name='FileLeafRef'/>" +
            $"<FieldRef Name='{sp.HubSpotDealIdColumn}'/>" +
            (hasOpportunityColumn ? $"<FieldRef Name='{sp.OpportunityIdColumn}'/>" : string.Empty) +
            $"<FieldRef Name='{sp.CustomerNameColumn}'/>" +
            $"<FieldRef Name='{sp.ProjectIdColumn}'/>" +
            $"<FieldRef Name='{sp.StatusColumn}'/></ViewFields>" +
            "<RowLimit Paged='TRUE'>500</RowLimit></View>",
        ListItemCollectionPosition = position,
    };

    var items = list.GetItems(query);
    if (hasOpportunityColumn)
    {
        ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(
            i => i["FileRef"], i => i["FileLeafRef"], i => i[sp.HubSpotDealIdColumn],
            i => i[sp.OpportunityIdColumn], i => i[sp.CustomerNameColumn],
            i => i[sp.ProjectIdColumn], i => i[sp.StatusColumn]));
    }
    else
    {
        ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(
            i => i["FileRef"], i => i["FileLeafRef"], i => i[sp.HubSpotDealIdColumn],
            i => i[sp.CustomerNameColumn], i => i[sp.ProjectIdColumn], i => i[sp.StatusColumn]));
    }

    await ctx.ExecuteQueryRetryAsync();

    foreach (var it in items)
    {
        var dealId = Field(it, sp.HubSpotDealIdColumn);
        if (string.IsNullOrWhiteSpace(dealId))
        {
            continue;
        }

        sets.Add((
            Field(it, "FileRef") ?? string.Empty,
            Field(it, "FileLeafRef") ?? string.Empty,
            dealId!,
            hasOpportunityColumn ? Field(it, sp.OpportunityIdColumn) : null,
            Field(it, sp.CustomerNameColumn),
            Field(it, sp.ProjectIdColumn),
            Field(it, sp.StatusColumn)));
    }

    position = items.ListItemCollectionPosition;
}
while (position is not null);

Console.WriteLine($"Found {sets.Count} workspace(s) carrying a HubSpot deal id.");
if (sets.Count == 0)
{
    Console.WriteLine("Nothing to do.");
    return 0;
}

var deals = await hubspot.GetDealsByIdAsync(sets.Select(s => s.DealId).ToList(), CancellationToken.None);
var opportunityByDeal = deals
    .Where(d => !string.IsNullOrWhiteSpace(d.OpportunityId))
    .ToDictionary(d => d.DealId, d => d.OpportunityId!.Trim(), StringComparer.OrdinalIgnoreCase);

Console.WriteLine($"Resolved an opportunity number for {opportunityByDeal.Count} of {sets.Count} deal(s).");
Console.WriteLine();

var toWrite = new List<(string Url, string DealId, string Value, string? Was)>();
var unresolved = new List<(string Url, string DealId)>();
var alreadyCorrect = 0;

foreach (var s in sets)
{
    if (!opportunityByDeal.TryGetValue(s.DealId, out var desired))
    {
        unresolved.Add((s.Url, s.DealId));
        continue;
    }

    if (string.Equals(s.Opportunity?.Trim(), desired, StringComparison.OrdinalIgnoreCase))
    {
        alreadyCorrect++;
        continue;
    }

    toWrite.Add((s.Url, s.DealId, desired, s.Opportunity));
}

// Folder renames: only for sets still in the scoping phase. A promoted set is named after its project
// id by design, and renaming it back to a PQCode form would undo that.
var toRename = new List<(string Url, string From, string To)>();
foreach (var s in sets)
{
    if (!string.IsNullOrWhiteSpace(s.ProjectId))
    {
        continue;
    }

    if (!opportunityByDeal.TryGetValue(s.DealId, out var pq))
    {
        continue;
    }

    var basis = string.IsNullOrWhiteSpace(s.Customer) ? pq : s.Customer!;
    var desiredLeaf = SharePointNaming.BuildDocumentSetName(basis, pq, sp.DocumentSetNameMaxLength);
    if (!string.Equals(s.Leaf, desiredLeaf, StringComparison.Ordinal))
    {
        toRename.Add((s.Url, s.Leaf, desiredLeaf));
    }
}

Console.WriteLine($"Already correct : {alreadyCorrect}");
Console.WriteLine($"To write        : {toWrite.Count}");
Console.WriteLine($"To rename       : {toRename.Count}");
Console.WriteLine($"Unresolved      : {unresolved.Count}");
Console.WriteLine();

if (toRename.Count > 0)
{
    Console.WriteLine(apply ? "Renaming:" : "Would rename:");
    foreach (var r in toRename)
    {
        Console.WriteLine($"    {r.From}");
        Console.WriteLine($"      → {r.To}");
    }
    Console.WriteLine();
}

if (toWrite.Count > 0)
{
    Console.WriteLine(apply ? "Writing:" : "Would write:");
    foreach (var w in toWrite)
    {
        var was = string.IsNullOrWhiteSpace(w.Was) ? "<blank>" : w.Was;
        Console.WriteLine($"    deal {w.DealId,-14} {was,-12} → {w.Value,-12} {w.Url}");
    }
    Console.WriteLine();
}

if (unresolved.Count > 0)
{
    // A deal HubSpot no longer returns, or one with no opportunity number set. These stay unmatchable
    // by PQCode — worth a look rather than a silent skip.
    Console.WriteLine("No opportunity number in HubSpot (left blank — PQCode will not match these):");
    foreach (var u in unresolved)
    {
        Console.WriteLine($"    deal {u.DealId,-14} {u.Url}");
    }
    Console.WriteLine();
}

if (!apply)
{
    Console.WriteLine("DRY RUN — nothing written. Re-run with --apply to write.");
    return 0;
}

var written = 0;
var failed = 0;
foreach (var w in toWrite)
{
    try
    {
        var item = ctx.Web.GetFolderByServerRelativeUrl(w.Url).ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();

        if (!item.FieldValues.ContainsKey(sp.OpportunityIdColumn))
        {
            Console.Error.WriteLine($"    ⚠ column '{sp.OpportunityIdColumn}' absent on {w.Url}; skipped.");
            failed++;
            continue;
        }

        item[sp.OpportunityIdColumn] = w.Value;
        item.Update();
        await ctx.ExecuteQueryRetryAsync();
        written++;
    }
    catch (Exception ex)
    {
        // Keep going: one bad item shouldn't abandon the rest of the backfill.
        Console.Error.WriteLine($"    ⚠ failed on {w.Url}: {ex.Message}");
        failed++;
    }
}

Console.WriteLine($"✔ Wrote {written} of {toWrite.Count}{(failed > 0 ? $"; {failed} failed" : "")}.");

// Renames come after the column writes: a rename changes the server-relative URL the writes address.
var renamed = 0;
foreach (var r in toRename)
{
    try
    {
        var item = ctx.Web.GetFolderByServerRelativeUrl(r.Url).ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();

        item["FileLeafRef"] = r.To;
        item.Update();
        await ctx.ExecuteQueryRetryAsync();
        renamed++;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"    ⚠ rename failed on {r.Url}: {ex.Message}");
        failed++;
    }
}

if (toRename.Count > 0)
{
    Console.WriteLine($"✔ Renamed {renamed} of {toRename.Count}.");
}

return failed > 0 ? 1 : 0;

static string? Field(ListItem item, string column)
    => item.FieldValues.TryGetValue(column, out var v) ? v?.ToString()?.Trim() : null;

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
