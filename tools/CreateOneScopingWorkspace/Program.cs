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
// Create ONE scoping workspace — a controlled live write for a single HubSpot deal.
// Resolves the deal's customer + owner (the production path) then creates/ensures
// the SharePoint document set (Status=Scoping) and reads back what landed.
//   Usage: dotnet run --project tools/CreateOneScopingWorkspace -- <dealId>
// -----------------------------------------------------------------------------

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: dotnet run --project tools/CreateOneScopingWorkspace -- <dealId>");
    return 1;
}

var dealId = args[0].Trim();

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var hubOptions = Bind<HubSpotOptions>(configuration, HubSpotOptions.SectionName);
var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => o.SingleLine = true));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(hubOptions.Value.TimeoutSeconds) };
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

var tokenProvider = new HubSpotTokenProvider(http, hubOptions, loggerFactory.CreateLogger<HubSpotTokenProvider>());
var hubspot = new HubSpotClient(http, tokenProvider, hubOptions, loggerFactory.CreateLogger<HubSpotClient>());
var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var uploadLinks = new GraphUploadLinkService(contextFactory, spOptions, loggerFactory.CreateLogger<GraphUploadLinkService>());
var sharePoint = new SharePointDocumentSetService(contextFactory, uploadLinks, spOptions, loggerFactory.CreateLogger<SharePointDocumentSetService>());

Console.WriteLine($"=== Create ONE scoping workspace: deal {dealId} ===");
Console.WriteLine();

// Find the deal (wide window) and resolve the same way the processor does.
var deals = await hubspot.GetDealsModifiedAfterAsync(DateTimeOffset.UtcNow.AddYears(-5), maxResults: 10000, CancellationToken.None);
var deal = deals.FirstOrDefault(d => d.DealId == dealId);
if (deal is null)
{
    Console.Error.WriteLine($"❌ Deal '{dealId}' not found (in scope: pipeline/terminal/modified filters). " +
        "It may be Won/Lost or outside the configured pipelines.");
    return 2;
}

var customer = await hubspot.ResolveCustomerNameAsync(deal, CancellationToken.None);
var owners = await hubspot.GetOwnerEmailsAsync(CancellationToken.None);
var ownerEmail = deal.OwnerId is { } oid && owners.TryGetValue(oid, out var em) ? em : null;

var siteUrl = string.IsNullOrWhiteSpace(spOptions.Value.PracticeMappings.First().SiteUrl)
    ? spOptions.Value.SiteUrl : spOptions.Value.PracticeMappings.First().SiteUrl!;

// Optional: recycle the existing scoping workspace for this deal (soft delete → Recycle Bin).
if (args.Contains("--delete"))
{
    using var delCtx = await contextFactory.CreateContextAsync(siteUrl);
    var delList = delCtx.Web.Lists.GetByTitle(spOptions.Value.PracticeMappings.First().Library);
    var col = spOptions.Value.HubSpotDealIdColumn;
    var query = new CamlQuery
    {
        ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>" +
                  $"<Eq><FieldRef Name='{col}'/><Value Type='Text'>{deal.DealId}</Value></Eq>" +
                  "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                  "</And></Where></Query><RowLimit>5</RowLimit></View>",
    };
    var items = delList.GetItems(query);
    delCtx.Load(items, i => i.Include(x => x["FileRef"]));
    await delCtx.ExecuteQueryRetryAsync();
    foreach (var it in items)
    {
        var url = it["FileRef"]?.ToString();
        Console.WriteLine($"Recycling: {url}");
        delCtx.Web.GetFolderByServerRelativeUrl(url).Recycle();
    }
    await delCtx.ExecuteQueryRetryAsync();
    Console.WriteLine(items.Count == 0 ? "No existing workspace found." : "✔ Recycled to the Recycle Bin (recoverable).");
    return 0;
}

Console.WriteLine("Deal:");
Console.WriteLine($"    Deal Id       : {deal.DealId}");
Console.WriteLine($"    Project name  : {deal.DealName}");
Console.WriteLine($"    Customer      : {customer}");
Console.WriteLine($"    Practice      : {deal.Practice}");
Console.WriteLine($"    Owner         : {ownerEmail ?? "<none>"}");
Console.WriteLine();

Console.WriteLine("Creating scoping workspace…");
var result = await sharePoint.EnsureScopingWorkspaceAsync(new ScopingWorkspace
{
    DealId = deal.DealId,
    CustomerName = customer,
    ProjectName = deal.DealName,
    Practice = deal.Practice,
    OwnerEmail = ownerEmail,
}, CancellationToken.None);

Console.WriteLine($"✔ {(result.Created ? "CREATED" : "already existed — updated")}: {result.ServerRelativeUrl}");
Console.WriteLine();

// Read back the metadata + permissions.
using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var folder = ctx.Web.GetFolderByServerRelativeUrl(result.ServerRelativeUrl);
var item = folder.ListItemAllFields;
ctx.Load(item);
ctx.Load(item, i => i.HasUniqueRoleAssignments);
var ras = item.RoleAssignments;
ctx.Load(ras, r => r.Include(a => a.Member.Title, a => a.RoleDefinitionBindings.Include(rd => rd.Name)));
await ctx.ExecuteQueryRetryAsync();

string Show(string col) => item.FieldValues.TryGetValue(col, out var v) && v is not null ? v.ToString()! : "<blank>";
Console.WriteLine("Metadata written:");
Console.WriteLine($"    {spOptions.Value.CustomerNameColumn,-22} = {Show(spOptions.Value.CustomerNameColumn)}");
Console.WriteLine($"    {spOptions.Value.ProjectNameColumn,-22} = {Show(spOptions.Value.ProjectNameColumn)}");
Console.WriteLine($"    {spOptions.Value.HubSpotDealIdColumn,-22} = {Show(spOptions.Value.HubSpotDealIdColumn)}");
Console.WriteLine($"    {spOptions.Value.StatusColumn,-22} = {Show(spOptions.Value.StatusColumn)}");
Console.WriteLine();
Console.WriteLine($"Permissions (inheritance broken: {item.HasUniqueRoleAssignments}):");
foreach (var ra in ras)
{
    Console.WriteLine($"    {ra.Member.Title,-32} : {string.Join(", ", ra.RoleDefinitionBindings.Select(rd => rd.Name))}");
}

if (spOptions.Value.CreateClientUploadLink)
{
    Console.WriteLine();
    Console.WriteLine("Client Uploads:");
    ctx.Load(folder.Folders, fs => fs.Include(f => f.Name, f => f.ServerRelativeUrl));
    await ctx.ExecuteQueryRetryAsync();
    var uploads = folder.Folders.FirstOrDefault(f => string.Equals(f.Name, spOptions.Value.ClientUploadsFolderName, StringComparison.OrdinalIgnoreCase));
    Console.WriteLine(uploads is not null ? $"    ✔ Subfolder : {uploads.ServerRelativeUrl}" : "    ⚠ Subfolder not found.");
    var linkCol = spOptions.Value.ClientUploadLinkColumn;
    if (item.FieldValues.TryGetValue(linkCol, out var lv) && lv is not null)
    {
        var url = lv is FieldUrlValue fu ? fu.Url : lv.ToString();
        Console.WriteLine($"    ✔ Link      : {url}");
    }
    else
    {
        Console.WriteLine($"    ⚠ Link column '{linkCol}' blank/absent.");
    }
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
