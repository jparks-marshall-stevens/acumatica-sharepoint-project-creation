using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// -----------------------------------------------------------------------------
// Rename a folder inside the configured library (site + library come from the
// first practice mapping). Lists the folder's contents first for safety.
//
// Usage: dotnet run --project tools/RenameFolder -- "<relPath>" "<newLeafName>" [--apply]
//   e.g. dotnet run --project tools/RenameFolder -- "Projects/Active" "Current" --apply
// Without --apply it only previews (lists contents, shows the intended rename).
// -----------------------------------------------------------------------------

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: RenameFolder -- \"<relPath>\" \"<newLeafName>\" [--apply]");
    return 1;
}

var relPath = args[0].Trim().Trim('/');
var newLeaf = args[1].Trim();
var apply = args.Contains("--apply");

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = new SharePointOptions();
configuration.GetSection(SharePointOptions.SectionName).Bind(spOptions);

var mapping = spOptions.PracticeMappings.FirstOrDefault()
    ?? throw new InvalidOperationException("No practice mapping configured.");
var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? spOptions.SiteUrl : mapping.SiteUrl!;
var library = mapping.Library;

using var loggerFactory = LoggerFactory.Create(b => b.SetMinimumLevel(LogLevel.Warning).AddSimpleConsole(o => o.SingleLine = true));
var factory = new SharePointContextFactory(
    Microsoft.Extensions.Options.Options.Create(spOptions),
    loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine("=== Rename folder ===");
Console.WriteLine($"Site    : {siteUrl}");
Console.WriteLine($"Library : {library}");
Console.WriteLine($"Folder  : {relPath}  ->  rename leaf to \"{newLeaf}\"");
Console.WriteLine($"Mode    : {(apply ? "APPLY" : "PREVIEW (pass --apply to execute)")}");
Console.WriteLine();

using var ctx = await factory.CreateContextAsync(siteUrl);
var list = ctx.Web.Lists.GetByTitle(library);
ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
await ctx.ExecuteQueryRetryAsync();

var folderUrl = $"{list.RootFolder.ServerRelativeUrl}/{relPath}";
var folder = ctx.Web.GetFolderByServerRelativeUrl(folderUrl);
ctx.Load(folder, f => f.Name, f => f.ServerRelativeUrl, f => f.Folders.Include(sf => sf.Name), f => f.Files.Include(fi => fi.Name));
var item = folder.ListItemAllFields;
ctx.Load(item);
await ctx.ExecuteQueryRetryAsync();

Console.WriteLine($"Found folder: {folder.ServerRelativeUrl}");
Console.WriteLine($"  Subfolders ({folder.Folders.Count}):");
foreach (var sf in folder.Folders) Console.WriteLine($"    - {sf.Name}");
Console.WriteLine($"  Files ({folder.Files.Count}):");
foreach (var fi in folder.Files) Console.WriteLine($"    - {fi.Name}");
Console.WriteLine();

if (!apply)
{
    Console.WriteLine($"PREVIEW only — would rename \"{folder.Name}\" to \"{newLeaf}\". Re-run with --apply to execute.");
    return 0;
}

item["FileLeafRef"] = newLeaf;
item.Update();
await ctx.ExecuteQueryRetryAsync();
Console.WriteLine($"✅ Renamed to \"{newLeaf}\". New path: {list.RootFolder.ServerRelativeUrl}/{Path.GetDirectoryName(relPath)?.Replace('\\','/')}/{newLeaf}".Replace("//", "/"));
return 0;

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!System.IO.File.Exists(path)) return result;
    using var doc = JsonDocument.Parse(System.IO.File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var values))
        foreach (var prop in values.EnumerateObject())
            result[prop.Name] = prop.Value.GetString();
    return result;
}
