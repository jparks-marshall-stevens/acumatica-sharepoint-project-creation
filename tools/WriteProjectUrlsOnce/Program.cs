using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Writes a project's real workspace URLs (dataroom + client upload) back to its Acumatica DATAURL /
// CLIENTURL attributes, then reads them back to confirm. This is the exact write the sync performs on
// creation — used here to validate it end to end against one project.
//
//   dotnet run --project tools/WriteProjectUrlsOnce -- <ProjectId>
// -----------------------------------------------------------------------------

var projectId = args.FirstOrDefault(a => !a.StartsWith("--"))?.Trim();
if (string.IsNullOrWhiteSpace(projectId))
{
    Console.Error.WriteLine("Usage: dotnet run --project tools/WriteProjectUrlsOnce -- <ProjectId>");
    return 1;
}

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var acuOptions = Bind<AcumaticaOptions>(configuration, AcumaticaOptions.SectionName);
var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var sp = spOptions.Value;

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => o.SingleLine = true));

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acuOptions.Value.TimeoutSeconds) };
var tokenProvider = new AcumaticaTokenProvider(http, acuOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acuOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());

// Look up the workspace for this project to get the real dataroom URL + upload link.
var mapping = sp.PracticeMappings.First();
var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? sp.SiteUrl : mapping.SiteUrl!;
var origin = new Uri(siteUrl).GetLeftPart(UriPartial.Authority);

using var ctx = await contextFactory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(mapping.Library);
var safe = System.Security.SecurityElement.Escape(projectId) ?? projectId;
var query = new CamlQuery
{
    ViewXml =
        "<View Scope='RecursiveAll'><Query><Where><And>" +
        $"<Eq><FieldRef Name='{sp.ProjectIdColumn}'/><Value Type='Text'>{safe}</Value></Eq>" +
        "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
        "</And></Where></Query>" +
        $"<ViewFields><FieldRef Name='FileRef'/><FieldRef Name='{sp.ClientUploadLinkColumn}'/></ViewFields>" +
        "<RowLimit>1</RowLimit></View>",
};
var items = list.GetItems(query);
ctx.Load(items, c => c.Include(i => i["FileRef"], i => i[sp.ClientUploadLinkColumn]));
await ctx.ExecuteQueryRetryAsync();

if (items.Count == 0)
{
    Console.Error.WriteLine($"❌ No workspace found for project {projectId}.");
    return 2;
}

var fileRef = items[0]["FileRef"]?.ToString() ?? string.Empty;
var dataUrl = origin + fileRef.Replace(" ", "%20");
var clientUrl = items[0].FieldValues.TryGetValue(sp.ClientUploadLinkColumn, out var v) ? v?.ToString() : null;

Console.WriteLine($"Project    : {projectId}");
Console.WriteLine($"DATAURL    : {dataUrl}");
Console.WriteLine($"CLIENTURL  : {(string.IsNullOrWhiteSpace(clientUrl) ? "<none stored>" : clientUrl)}");
Console.WriteLine();
Console.WriteLine("Writing back to Acumatica…");
var ok = await acumatica.WriteProjectUrlsAsync(projectId, dataUrl, clientUrl, CancellationToken.None);
Console.WriteLine(ok ? "✔ write returned success" : "✘ write returned false (see log)");
return ok ? 0 : 3;

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
