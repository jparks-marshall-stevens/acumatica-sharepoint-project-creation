using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Backfill missing client-upload links. The sync mints a room's anonymous "Request files" link only once,
// when the room is created, and never retries. Rooms created on a site whose sharing didn't allow Anyone
// links (the new practice sites until 2026-09-28) therefore have an empty ClientUploadLink.
//
// For each document set on the given site(s) with no ClientUploadLink this tool:
//   1. ensures the "Client Uploads" subfolder exists,
//   2. mints the upload link via Graph (same call and settings as the sync),
//   3. stamps it in ClientUploadLink,
//   4. writes it to the Acumatica project's CLIENTURL attribute (execution rooms only; DATAURL is untouched).
//
//   dotnet run --project tools/BackfillUploadLinks -- <siteSlug> [<siteSlug> ...] [--apply] [--limit N]
//   e.g. dotnet run --project tools/BackfillUploadLinks -- TransactionAdvisory ESOP
// Without --apply it only lists the rooms that would be fixed. --limit caps how many rooms are processed
// (handy for trying one room first).
// -----------------------------------------------------------------------------

var apply = args.Any(a => a.Equals("--apply", StringComparison.OrdinalIgnoreCase));
var limit = int.MaxValue;
var slugs = new List<string>();
for (var i = 0; i < args.Length; i++)
{
    if (args[i].Equals("--apply", StringComparison.OrdinalIgnoreCase)) continue;
    if (args[i].Equals("--limit", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var n))
    {
        limit = n;
        i++;
        continue;
    }
    slugs.Add(args[i].Trim().Trim('/'));
}

if (slugs.Count == 0)
{
    Console.Error.WriteLine("Usage: BackfillUploadLinks -- <siteSlug> [<siteSlug> ...] [--apply] [--limit N]");
    return 1;
}

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
    .AddSimpleConsole(o => o.SingleLine = true));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acuOptions.Value.TimeoutSeconds) };
var tokenProvider = new AcumaticaTokenProvider(http, acuOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acuOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var uploadLinks = new GraphUploadLinkService(contextFactory, spOptions, loggerFactory.CreateLogger<GraphUploadLinkService>());

var origin = new Uri(string.IsNullOrWhiteSpace(sp.SiteUrl) ? sp.PracticeMappings.First().SiteUrl! : sp.SiteUrl)
    .GetLeftPart(UriPartial.Authority);
var library = sp.PracticeMappings.FirstOrDefault()?.Library ?? "Documents";
var pidCol = sp.ProjectIdColumn;
var linkCol = sp.ClientUploadLinkColumn;
var uploadsName = SharePointNaming.SanitizeLeafName(sp.ClientUploadsFolderName);
var writeBack = !string.IsNullOrWhiteSpace(acuOptions.Value.ClientUrlAttributeId);

Console.WriteLine($"Mode      : {(apply ? "APPLY" : "PREVIEW (pass --apply to execute)")}");
Console.WriteLine($"Sites     : {string.Join(", ", slugs)}");
Console.WriteLine($"Library   : {library}");
Console.WriteLine($"Link      : {sp.ClientUploadLinkScope}, expires in {sp.ClientUploadLinkExpirationDays} days");
Console.WriteLine($"Acumatica : {(writeBack ? $"write {acuOptions.Value.ClientUrlAttributeId} for execution rooms" : "write-back off (no ClientUrlAttributeId)")}");
Console.WriteLine();

int found = 0, linked = 0, failed = 0, acuWritten = 0, acuFailed = 0;
var acumaticaStopped = false;

foreach (var slug in slugs)
{
    var siteUrl = $"{origin}/sites/{slug}";
    Console.WriteLine($"=== {siteUrl}");

    using var ctx = await contextFactory.CreateContextAsync(siteUrl);
    var list = ctx.Web.Lists.GetByTitle(library);
    ctx.Load(list, l => l.RootFolder.ServerRelativeUrl, l => l.Fields.Include(f => f.InternalName));
    await ctx.ExecuteQueryRetryAsync();

    if (!list.Fields.Any(f => f.InternalName == linkCol))
    {
        if (!apply)
        {
            Console.WriteLine($"  (column '{linkCol}' is missing; --apply will create it)");
        }
        else
        {
            // Same shape the sync creates: plain Text, so the value shows as a copyable URL.
            list.Fields.AddFieldAsXml(
                $"<Field Type='Text' Name='{linkCol}' StaticName='{linkCol}' DisplayName='{linkCol}' Group='ProjectSync'/>",
                addToDefaultView: false, options: AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryRetryAsync();
            Console.WriteLine($"  Created column '{linkCol}'.");
        }
    }
    var hasLinkCol = list.Fields.Any(f => f.InternalName == linkCol);

    // Every document set (Document Set content types all start with 0x0120D520) with no upload link.
    var rooms = new List<(string FileRef, string? ProjectId)>();
    ListItemCollectionPosition? position = null;
    do
    {
        var query = new CamlQuery
        {
            ViewXml =
                "<View Scope='RecursiveAll'><Query><Where><And>" +
                "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                "<BeginsWith><FieldRef Name='ContentTypeId'/><Value Type='ContentTypeId'>0x0120D520</Value></BeginsWith>" +
                "</And></Where></Query>" +
                $"<ViewFields><FieldRef Name='FileRef'/><FieldRef Name='{pidCol}'/>" +
                (hasLinkCol ? $"<FieldRef Name='{linkCol}'/>" : "") +
                "</ViewFields><RowLimit Paged='TRUE'>2000</RowLimit></View>",
            ListItemCollectionPosition = position,
        };
        var items = list.GetItems(query);
        ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(i => i["FileRef"], i => i[pidCol]));
        if (hasLinkCol) ctx.Load(items, c => c.Include(i => i[linkCol]));
        await ctx.ExecuteQueryRetryAsync();

        foreach (var it in items)
        {
            var link = hasLinkCol && it.FieldValues.TryGetValue(linkCol, out var lv) ? lv?.ToString() : null;
            if (!string.IsNullOrWhiteSpace(link)) continue;
            var pid = it.FieldValues.TryGetValue(pidCol, out var pv) ? pv?.ToString()?.Trim() : null;
            rooms.Add((it["FileRef"]?.ToString() ?? "", string.IsNullOrWhiteSpace(pid) ? null : pid));
        }
        position = items.ListItemCollectionPosition;
    }
    while (position is not null);

    Console.WriteLine($"  {rooms.Count} room(s) without an upload link.");

    foreach (var (fileRef, pid) in rooms.OrderBy(r => r.FileRef, StringComparer.OrdinalIgnoreCase))
    {
        if (found >= limit) break;
        found++;
        var label = $"{Path.GetFileName(fileRef)}{(pid is null ? "  (scoping)" : $"  [{pid}]")}";

        if (!apply)
        {
            Console.WriteLine($"  - {label}");
            continue;
        }

        string? newLink;
        try
        {
            // 1) Client Uploads subfolder (the sync normally created it already).
            var uploadsUrl = $"{fileRef.TrimEnd('/')}/{uploadsName}";
            var uploads = ctx.Web.GetFolderByServerRelativeUrl(uploadsUrl);
            ctx.Load(uploads, f => f.Exists);
            await ctx.ExecuteQueryRetryAsync();
            if (!uploads.Exists)
            {
                ctx.Web.GetFolderByServerRelativeUrl(fileRef).Folders.Add(uploadsName);
                await ctx.ExecuteQueryRetryAsync();
            }

            // 2) Mint the link (logs the Graph error body itself on failure).
            newLink = await uploadLinks.CreateUploadLinkAsync(siteUrl, list.RootFolder.ServerRelativeUrl, uploadsUrl, CancellationToken.None);
            if (string.IsNullOrWhiteSpace(newLink))
            {
                failed++;
                Console.WriteLine($"  ✘ {label}: no link (is the site's sharing set to Anyone?)");
                continue;
            }

            // 3) Stamp it on the document set item.
            var item = ctx.Web.GetFolderByServerRelativeUrl(fileRef).ListItemAllFields;
            item[linkCol] = newLink;
            item.Update();
            await ctx.ExecuteQueryRetryAsync();
            linked++;
        }
        catch (Exception ex)
        {
            failed++;
            Console.WriteLine($"  ✘ {label}: {ex.Message}");
            continue;
        }

        // 4) Acumatica CLIENTURL. Stop writing on the first exception (e.g. a token failure) rather than
        //    retrying login once per room and risking a lockout of the service account.
        var acu = "";
        if (pid is not null && writeBack && !acumaticaStopped)
        {
            try
            {
                if (await acumatica.WriteProjectUrlsAsync(pid, dataUrl: null, clientUrl: newLink, CancellationToken.None))
                {
                    acuWritten++;
                    acu = "  + Acumatica";
                }
                else
                {
                    acuFailed++;
                    acu = "  (Acumatica write failed)";
                }
            }
            catch (Exception ex)
            {
                acuFailed++;
                acumaticaStopped = true;
                acu = $"  (Acumatica error, no more Acumatica writes this run: {ex.Message})";
            }
        }
        Console.WriteLine($"  ✔ {label}{acu}");
    }

    Console.WriteLine();
    if (found >= limit) break;
}

Console.WriteLine(apply
    ? $"Done: {linked} link(s) created, {failed} failed; Acumatica CLIENTURL written {acuWritten}, failed {acuFailed}."
    : $"PREVIEW: {found} room(s) would get a link. Re-run with --apply to execute.");
return failed > 0 || acuFailed > 0 ? 3 : 0;

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
