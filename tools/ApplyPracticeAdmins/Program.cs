using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using System.Text.Json;

// -----------------------------------------------------------------------------
// Apply practice administrators to EXISTING document sets.
//
// The sync grants each practice's AdminEmails (SharePoint:PracticeMappings:N:AdminEmails) on every
// workspace it touches — but a workspace is only re-permissioned on its own cadence (execution sets on
// the daily reconcile; scoping sets when their deal next changes). This tool bridges that gap: it grants
// the configured admins access to every existing workspace NOW, so a newly-added admin doesn't have to
// wait for each folder to re-sync. It is ADDITIVE — it grants the admin the configured PermissionLevel
// and touches nothing else about a folder's access. Re-runnable and idempotent.
//
//   dotnet run --project tools/ApplyPracticeAdmins              → dry run (default; writes nothing)
//   dotnet run --project tools/ApplyPracticeAdmins -- --apply   → grant access
// -----------------------------------------------------------------------------

var apply = args.Contains("--apply", StringComparer.OrdinalIgnoreCase);

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);
var sp = spOptions.Value;

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information)
    .AddSimpleConsole(o => o.SingleLine = true));

var contextFactory = new SharePointContextFactory(spOptions, loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine($"=== Apply practice admins to existing workspaces{(apply ? "" : "  (DRY RUN)")} ===");
Console.WriteLine();

// Group practices by (site, library) so each library is swept once even if several practices share it.
var groups = sp.PracticeMappings
    .Where(m => m.AdminEmails is { Count: > 0 })
    .GroupBy(m => (
        Site: string.IsNullOrWhiteSpace(m.SiteUrl) ? sp.SiteUrl : m.SiteUrl!,
        m.Library))
    .ToList();

if (groups.Count == 0)
{
    Console.WriteLine("No practice mapping has AdminEmails configured. Nothing to do.");
    return 0;
}

var totalGranted = 0;
var totalAlready = 0;
var totalFailed = 0;

foreach (var group in groups)
{
    var admins = group.SelectMany(m => m.AdminEmails)
        .Select(a => a.Trim())
        .Where(a => a.Length > 0)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    Console.WriteLine($"Site    : {group.Key.Site}");
    Console.WriteLine($"Library : {group.Key.Library}");
    Console.WriteLine($"Admins  : {string.Join(", ", admins)}");
    Console.WriteLine($"Level   : {sp.PermissionLevel}");
    Console.WriteLine();

    using var ctx = await contextFactory.CreateContextAsync(group.Key.Site);
    var list = ctx.Web.Lists.GetByTitle(group.Key.Library);
    var role = ctx.Web.RoleDefinitions.GetByName(sp.PermissionLevel);
    ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
    ctx.Load(role, r => r.Id, r => r.Name);
    await ctx.ExecuteQueryRetryAsync();

    // Resolve each admin to a site user once (fail-soft: an unresolvable email is reported and skipped).
    var adminUsers = new List<User>();
    foreach (var email in admins)
    {
        try
        {
            var u = ctx.Web.EnsureUser(email);
            ctx.Load(u, x => x.Id, x => x.Title, x => x.LoginName);
            await ctx.ExecuteQueryRetryAsync();
            adminUsers.Add(u);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"    ⚠ could not resolve '{email}': {ex.Message}");
            totalFailed++;
        }
    }

    if (adminUsers.Count == 0)
    {
        Console.WriteLine("  No resolvable admins for this library; skipping.");
        Console.WriteLine();
        continue;
    }

    // Every document set (folder) in the library.
    var docSets = new List<string>();
    ListItemCollectionPosition? position = null;
    do
    {
        var query = new CamlQuery
        {
            ViewXml =
                "<View Scope='RecursiveAll'><Query><Where>" +
                "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                "</Where></Query>" +
                "<ViewFields><FieldRef Name='FileRef'/><FieldRef Name='ContentTypeId'/></ViewFields>" +
                "<RowLimit Paged='TRUE'>500</RowLimit></View>",
            ListItemCollectionPosition = position,
        };
        var items = list.GetItems(query);
        ctx.Load(items, c => c.ListItemCollectionPosition,
            c => c.Include(i => i["FileRef"], i => i["ContentTypeId"]));
        await ctx.ExecuteQueryRetryAsync();

        foreach (var it in items)
        {
            // Only Document Sets (the "Project" content type), not incidental sub-folders inside them.
            var ct = it["ContentTypeId"]?.ToString() ?? string.Empty;
            var url = it["FileRef"]?.ToString();
            if (!string.IsNullOrWhiteSpace(url) && ct.StartsWith("0x0120D520", StringComparison.OrdinalIgnoreCase))
            {
                docSets.Add(url!);
            }
        }

        position = items.ListItemCollectionPosition;
    }
    while (position is not null);

    Console.WriteLine($"  {docSets.Count} document set(s) found.");
    Console.WriteLine();

    foreach (var url in docSets)
    {
        try
        {
            var item = ctx.Web.GetFolderByServerRelativeUrl(url).ListItemAllFields;
            ctx.Load(item, i => i.HasUniqueRoleAssignments,
                i => i.RoleAssignments.Include(ra => ra.PrincipalId));
            await ctx.ExecuteQueryRetryAsync();

            var present = item.RoleAssignments.Select(ra => ra.PrincipalId).ToHashSet();
            var missing = adminUsers.Where(u => !present.Contains(u.Id)).ToList();

            if (missing.Count == 0)
            {
                totalAlready++;
                continue;
            }

            var who = string.Join(", ", missing.Select(u => u.Title));
            Console.WriteLine($"    {(apply ? "grant" : "would grant")} [{who}] on {url}");

            if (apply)
            {
                // The authoritative sync already breaks inheritance on these sets; if one somehow still
                // inherits, break it while PRESERVING current access so we only add, never remove.
                if (!item.HasUniqueRoleAssignments)
                {
                    item.BreakRoleInheritance(copyRoleAssignments: true, clearSubscopes: false);
                }

                foreach (var u in missing)
                {
                    var binding = new RoleDefinitionBindingCollection(ctx) { role };
                    item.RoleAssignments.Add(u, binding);
                }

                await ctx.ExecuteQueryRetryAsync();
            }

            totalGranted++;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"    ⚠ failed on {url}: {ex.Message}");
            totalFailed++;
        }
    }

    Console.WriteLine();
}

Console.WriteLine($"{(apply ? "Granted" : "Would grant")} on {totalGranted} set(s); already had access: {totalAlready}; failed: {totalFailed}.");
if (!apply)
{
    Console.WriteLine("DRY RUN — nothing written. Re-run with --apply to grant.");
}

return totalFailed > 0 ? 1 : 0;

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
