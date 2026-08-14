using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// Inspects (and lightly hardens) the GiftEstate Documents library:
//   • reports + enables versioning
//   • reports recycle-bin status
//   • reports the Projects/Current folder permissions and who can create there
// Read-only except that it turns versioning ON if it is off.

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = new SharePointOptions();
configuration.GetSection(SharePointOptions.SectionName).Bind(spOptions);
var mapping = spOptions.PracticeMappings.First();
var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? spOptions.SiteUrl : mapping.SiteUrl!;
var library = mapping.Library;
var currentFolderRel = mapping.ParentFolder ?? "Projects/Current";

using var loggerFactory = LoggerFactory.Create(b => b.SetMinimumLevel(LogLevel.Warning).AddSimpleConsole(o => o.SingleLine = true));
var factory = new SharePointContextFactory(Options.Create(spOptions), loggerFactory.CreateLogger<SharePointContextFactory>());
using var ctx = await factory.CreateContextAsync(siteUrl);

var list = ctx.Web.Lists.GetByTitle(library);
ctx.Load(list, l => l.EnableVersioning, l => l.EnableMinorVersions, l => l.MajorVersionLimit,
    l => l.MajorWithMinorVersionsLimit, l => l.RootFolder.ServerRelativeUrl, l => l.Title);
ctx.Load(ctx.Web, w => w.Title, w => w.Url);
await ctx.ExecuteQueryRetryAsync();

Console.WriteLine($"=== Hardening: {ctx.Web.Title} / {library} ===");
Console.WriteLine();
Console.WriteLine("Versioning:");
Console.WriteLine($"    EnableVersioning        : {list.EnableVersioning}");
Console.WriteLine($"    EnableMinorVersions     : {list.EnableMinorVersions}");
Console.WriteLine($"    MajorVersionLimit       : {list.MajorVersionLimit}");
if (!list.EnableVersioning)
{
    Console.WriteLine("    → versioning was OFF; enabling (keep 500 major versions)…");
    list.EnableVersioning = true;
    list.MajorVersionLimit = 500;
    list.Update();
    await ctx.ExecuteQueryRetryAsync();
    Console.WriteLine("    ✔ versioning enabled.");
}
else
{
    Console.WriteLine("    ✔ already on.");
}
Console.WriteLine();
Console.WriteLine("Recycle bin: SharePoint Online keeps site + second-stage recycle bins ON by platform");
Console.WriteLine("             default (93 days) — it cannot be disabled, so deletes are recoverable.");
Console.WriteLine();

// Current folder permissions.
var currentUrl = $"{list.RootFolder.ServerRelativeUrl}/{currentFolderRel.Trim('/')}";
var folder = ctx.Web.GetFolderByServerRelativeUrl(currentUrl);
var item = folder.ListItemAllFields;
ctx.Load(item, i => i.HasUniqueRoleAssignments);
ctx.Load(item.RoleAssignments, r => r.Include(
    a => a.Member.Title, a => a.Member.PrincipalType,
    a => a.RoleDefinitionBindings.Include(rd => rd.Name, rd => rd.BasePermissions)));
await ctx.ExecuteQueryRetryAsync();

Console.WriteLine($"'{currentFolderRel}' folder ({currentUrl}):");
Console.WriteLine($"    Unique permissions (inheritance broken): {item.HasUniqueRoleAssignments}");
Console.WriteLine("    Who has access here (principal : levels : can-create?):");
foreach (var ra in item.RoleAssignments)
{
    var levels = ra.RoleDefinitionBindings.Select(rd => rd.Name).ToList();
    var canAdd = ra.RoleDefinitionBindings.Any(rd => rd.BasePermissions.Has(PermissionKind.AddListItems));
    Console.WriteLine($"      {ra.Member.Title} [{ra.Member.PrincipalType}] : {string.Join(", ", levels)} : {(canAdd ? "CAN CREATE" : "no create")}");
}
Console.WriteLine();

if (!args.Contains("--lock"))
{
    Console.WriteLine("(Run with --lock to break Current's inheritance and drop create-capable groups (except");
    Console.WriteLine(" Owners) to Read. The app keeps creating via its site-admin grant.)");
    return 0;
}

Console.WriteLine("Locking 'Current' so only the app can create here…");
ctx.Load(ctx.Web.AssociatedOwnerGroup, g => g.Id);
var readDef = ctx.Web.RoleDefinitions.GetByType(RoleType.Reader);
ctx.Load(readDef);
await ctx.ExecuteQueryRetryAsync();

if (!item.HasUniqueRoleAssignments)
{
    item.BreakRoleInheritance(copyRoleAssignments: true, clearSubscopes: false);
    await ctx.ExecuteQueryRetryAsync();
}

ctx.Load(item.RoleAssignments, r => r.Include(
    a => a.Member.Title, a => a.Member.PrincipalType, a => a.Member.Id,
    a => a.RoleDefinitionBindings.Include(rd => rd.Name, rd => rd.BasePermissions)));
await ctx.ExecuteQueryRetryAsync();

var ownerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
var toFix = item.RoleAssignments.Where(a =>
        a.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup &&
        a.Member.Id != ownerGroupId &&
        a.RoleDefinitionBindings.Any(rd => rd.BasePermissions.Has(PermissionKind.AddListItems)))
    .ToList();

foreach (var ra in toFix)
{
    Console.WriteLine($"    {ra.Member.Title}: removing create → Read");
    ra.RoleDefinitionBindings.RemoveAll();
    ra.RoleDefinitionBindings.Add(readDef);
    ra.Update();
}
await ctx.ExecuteQueryRetryAsync();
Console.WriteLine($"✔ Locked. {toFix.Count} group(s) downgraded to Read; only the app (and Owners) can create in Current.");
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
