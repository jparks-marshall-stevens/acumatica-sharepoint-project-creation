using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// -----------------------------------------------------------------------------
// Provision a new practice SharePoint site so it MIRRORS the Gift & Estate site
// (the "source"): the custom "Project" Document Set content type, the four base
// metadata columns, the Projects/Current folder scaffold (locked to code-only
// creation), and library versioning.
//
// This is the REPEATABLE step. It runs app-only with the SAME certificate as the
// sync (like every other tool here), so the target site MUST already exist and the
// cert app MUST already hold a Sites.Selected `fullControl` grant on it. Those two
// one-time, admin-only steps are done from bootstrap-practice-site.ps1 (no app
// registration required) — this tool prints exactly how if it can't reach the target.
//
// The four base columns' internal names are GLOBAL config (SharePoint:*Column) shared
// by every site the sync writes to, so a new site cannot use fresh clean names — it
// must reuse Project_x0020_Id, Customer_x0020_Name, Project_x0020_Name, and the People
// field Project_x0020_Manager. This tool reads them straight off the source site to
// guarantee an exact match. The runtime auto-creates the remaining text columns
// (Status, OpportunityId, DataroomUrl, ClientUploadLink, ProjectSyncSig) on first write.
//
// Usage:
//   dotnet run --project tools/ProvisionPracticeSite -- \
//       --practice "Marital Dissolution" \
//       --leader   rhoffman@marshall-stevens.com \
//       --to       https://marshallstevens.sharepoint.com/sites/MaritalDissolution \
//       [--from    https://marshallstevens.sharepoint.com/sites/GiftEstate]  (default: SharePoint:SiteUrl)
//       [--library Documents] [--parent-folder Projects/Current]
//       [--no-lock]   (skip locking the Current folder to code-only creation)
//       [--apply]     (default is a dry run — nothing is written)
//
// Dry run by default; idempotent; safe to re-run.
// -----------------------------------------------------------------------------

const string DocSetBaseId = "0x0120D520"; // any content-type id under this is a Document Set type

var opts = ParseArgs(args);
if (opts is null)
{
    return 1;
}

var apply = opts.Apply;

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var spOptions = new SharePointOptions();
configuration.GetSection(SharePointOptions.SectionName).Bind(spOptions);

var fromSite = opts.FromSite ?? spOptions.SiteUrl;
var library = opts.Library ?? spOptions.PracticeMappings.FirstOrDefault()?.Library ?? "Documents";
var parentFolder = opts.ParentFolder ?? spOptions.PracticeMappings.FirstOrDefault()?.ParentFolder ?? "Projects/Current";

// The four base columns the sync writes and the CT must carry (PracticeColumn is intentionally blank).
var baseColumns = new[]
{
    spOptions.ProjectIdColumn,
    spOptions.CustomerNameColumn,
    spOptions.ProjectNameColumn,
    spOptions.ProjectManagerColumn,
}.Where(c => !string.IsNullOrWhiteSpace(c)).Distinct().ToArray();

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddSimpleConsole(o => o.SingleLine = true));
var factory = new SharePointContextFactory(Options.Create(spOptions), loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine($"=== Provision practice site{(apply ? "" : "  (DRY RUN — nothing written)")} ===");
Console.WriteLine($"Practice   : {opts.Practice}");
Console.WriteLine($"Leader     : {opts.Leader ?? "(none)"}");
Console.WriteLine($"Source     : {fromSite}");
Console.WriteLine($"Target     : {opts.ToSite}");
Console.WriteLine($"Library    : {library}");
Console.WriteLine($"Parent     : {parentFolder}");
Console.WriteLine($"Content type: {spOptions.DocumentSetContentType}");
Console.WriteLine($"Base columns: {string.Join(", ", baseColumns)}");
Console.WriteLine();

// -----------------------------------------------------------------------------
// 1. Read the mirror source (Gift & Estate).
// -----------------------------------------------------------------------------
SourceModel source;
try
{
    source = await ReadSourceAsync(factory, fromSite, library, spOptions.DocumentSetContentType, baseColumns);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Could not read the source site {fromSite}: {ex.Message}");
    return 2;
}

Console.WriteLine("Source read OK:");
Console.WriteLine($"  Web template     : {source.WebTemplate}");
Console.WriteLine($"  '{spOptions.DocumentSetContentType}' CT parent id : {source.ParentContentTypeId} ({(source.ParentContentTypeId.StartsWith(DocSetBaseId, StringComparison.OrdinalIgnoreCase) ? "Document Set" : "NOT a Document Set!")})");
Console.WriteLine($"  Versioning       : Enable={source.EnableVersioning}, MajorLimit={source.MajorVersionLimit}");
foreach (var f in source.Fields)
{
    Console.WriteLine($"  Column           : {f.InternalName} ({f.Type})");
}
Console.WriteLine();

var missingOnSource = baseColumns.Where(c => source.Fields.All(f => f.InternalName != c)).ToArray();
if (missingOnSource.Length > 0)
{
    Console.Error.WriteLine($"❌ Source site is missing configured column(s): {string.Join(", ", missingOnSource)}. Aborting — cannot mirror what isn't there.");
    return 2;
}
if (!source.ParentContentTypeId.StartsWith(DocSetBaseId, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"❌ Source content type '{spOptions.DocumentSetContentType}' does not derive from a Document Set. Aborting.");
    return 2;
}

// -----------------------------------------------------------------------------
// 2. Connect to the target and confirm the one-time bootstrap has been done.
// -----------------------------------------------------------------------------
ClientContext ctx;
try
{
    ctx = await factory.CreateContextAsync(opts.ToSite);
    ctx.Load(ctx.Web, w => w.Title, w => w.Url, w => w.AssociatedOwnerGroup);
    await ctx.ExecuteQueryRetryAsync();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Could not reach the target site {opts.ToSite}: {ex.Message}");
    Console.Error.WriteLine();
    PrintBootstrapHelp(opts.ToSite, opts.Title ?? opts.Practice, fromSite, spOptions.ClientId, spOptions.AzureAdTenant);
    return 3;
}

Console.WriteLine($"✔ Reached target web \"{ctx.Web.Title}\" ({ctx.Web.Url}). The cert grant is in place.");
Console.WriteLine();

using (ctx)
{
    var list = ctx.Web.Lists.GetByTitle(library);
    ctx.Load(list, l => l.Title, l => l.ContentTypesEnabled, l => l.EnableVersioning,
        l => l.MajorVersionLimit, l => l.RootFolder.ServerRelativeUrl);
    ctx.Load(list.ContentTypes, cts => cts.Include(c => c.Name, c => c.StringId));
    ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName));
    await ctx.ExecuteQueryRetryAsync();

    // --- 3. Enable content types on the library ---
    if (!list.ContentTypesEnabled)
    {
        Console.WriteLine($"{Tag(apply)} enable content types on '{library}'.");
        if (apply)
        {
            list.ContentTypesEnabled = true;
            list.Update();
            await ctx.ExecuteQueryRetryAsync();
        }
    }
    else
    {
        Console.WriteLine($"✔ content types already enabled on '{library}'.");
    }

    // --- 4. Create the four base columns as SITE columns, mirroring source SchemaXml ---
    // Site columns (web.Fields) so they can be bound to a site content type below.
    ctx.Load(ctx.Web.Fields, fs => fs.Include(f => f.InternalName));
    await ctx.ExecuteQueryRetryAsync();
    var existingWebFieldNames = new HashSet<string>(ctx.Web.Fields.Select(f => f.InternalName), StringComparer.Ordinal);

    foreach (var srcField in source.Fields)
    {
        if (existingWebFieldNames.Contains(srcField.InternalName))
        {
            Console.WriteLine($"✔ column '{srcField.InternalName}' already exists.");
            continue;
        }

        Console.WriteLine($"{Tag(apply)} create column '{srcField.InternalName}' ({srcField.Type}) from source schema.");
        if (apply)
        {
            ctx.Web.Fields.AddFieldAsXml(srcField.SchemaXml, addToDefaultView: false,
                options: AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryRetryAsync();
        }
    }

    // --- 5. Create the "Project" Document Set content type + bind the columns ---
    var projectCt = list.ContentTypes.FirstOrDefault(c =>
        string.Equals(c.Name, spOptions.DocumentSetContentType, StringComparison.OrdinalIgnoreCase));

    if (projectCt is not null)
    {
        var isDocSet = projectCt.StringId.StartsWith(DocSetBaseId, StringComparison.OrdinalIgnoreCase);
        Console.WriteLine($"✔ content type '{spOptions.DocumentSetContentType}' already on '{library}' ({(isDocSet ? "Document Set" : "⚠ NOT a Document Set")}).");
    }
    else if (apply)
    {
        // Derive from the BUILT-IN Document Set content type (0x0120D520), NOT from the source's parent id:
        // Gift & Estate's "Project" derives from a G&E-specific custom Document Set CT that does not exist
        // on a fresh site. Deriving from the OOTB base gives a real, fully-wired Document Set CT here (a
        // new id under 0x0120D520 — invisible to the sync, which only checks the prefix + the columns).
        // The OOTB CT only exists once the "Document Sets" site-collection feature is active, so ensure it.
        await EnsureDocumentSetFeatureAsync(ctx);
        var parent = ctx.Web.ContentTypes.GetById(DocSetBaseId);
        ctx.Load(parent, c => c.Id, c => c.Name);
        await ctx.ExecuteQueryRetryAsync();

        var ci = new ContentTypeCreationInformation
        {
            Name = spOptions.DocumentSetContentType,
            Group = "Marshall & Stevens",
            ParentContentType = parent,
        };
        var newCt = ctx.Web.ContentTypes.Add(ci);
        ctx.Load(newCt, c => c.Id, c => c.StringId, c => c.Name);
        await ctx.ExecuteQueryRetryAsync();
        Console.WriteLine($"  created site content type '{newCt.Name}' ({newCt.StringId}).");

        // Bind the base columns as field links.
        ctx.Load(ctx.Web.Fields, fs => fs.Include(f => f.InternalName, f => f.Id));
        await ctx.ExecuteQueryRetryAsync();
        foreach (var srcField in source.Fields)
        {
            var field = ctx.Web.Fields.First(f => f.InternalName == srcField.InternalName);
            newCt.FieldLinks.Add(new FieldLinkCreationInformation { Field = field });
        }
        newCt.Update(updateChildren: true);
        await ctx.ExecuteQueryRetryAsync();
        Console.WriteLine($"  bound {source.Fields.Count} column(s) to '{newCt.Name}'.");

        // Add the CT to the library.
        list.ContentTypes.AddExistingContentType(newCt);
        await ctx.ExecuteQueryRetryAsync();
        Console.WriteLine($"  added '{newCt.Name}' to library '{library}'.");
    }
    else
    {
        Console.WriteLine($"{Tag(apply)} create Document Set content type '{spOptions.DocumentSetContentType}' (derive from built-in Document Set {DocSetBaseId}), bind {source.Fields.Count} column(s), add to '{library}'.");
    }

    // --- 6. Versioning (mirror source) ---
    if (list.EnableVersioning != source.EnableVersioning ||
        (source.EnableVersioning && list.MajorVersionLimit != source.MajorVersionLimit))
    {
        Console.WriteLine($"{Tag(apply)} set versioning: Enable={source.EnableVersioning}, MajorLimit={source.MajorVersionLimit}.");
        if (apply)
        {
            list.EnableVersioning = source.EnableVersioning;
            if (source.EnableVersioning && source.MajorVersionLimit > 0)
            {
                list.MajorVersionLimit = source.MajorVersionLimit;
            }
            list.Update();
            await ctx.ExecuteQueryRetryAsync();
        }
    }
    else
    {
        Console.WriteLine($"✔ versioning already matches source (Enable={list.EnableVersioning}, MajorLimit={list.MajorVersionLimit}).");
    }

    // --- 7. Folder scaffold (Projects/Current) ---
    var scaffold = parentFolder.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
    var currentUrl = $"{list.RootFolder.ServerRelativeUrl}/{string.Join('/', scaffold)}";
    if (apply)
    {
        var folder = list.RootFolder;
        foreach (var seg in scaffold)
        {
            folder = folder.Folders.Add(SharePointNaming.SanitizeLeafName(seg));
            ctx.Load(folder);
        }
        await ctx.ExecuteQueryRetryAsync();
        Console.WriteLine($"✔ ensured folder scaffold '{parentFolder}'.");
    }
    else
    {
        Console.WriteLine($"{Tag(apply)} ensure folder scaffold '{parentFolder}'.");
    }

    // --- 8. Lock 'Current' to code-only creation (mirror SharePointHardening --lock) ---
    if (opts.Lock)
    {
        if (apply)
        {
            await LockFolderAsync(ctx, currentUrl);
        }
        else
        {
            Console.WriteLine($"{Tag(apply)} lock '{parentFolder}': break inheritance, drop create-capable groups (except Owners) to Read.");
        }
    }
    else
    {
        Console.WriteLine("(skipping folder lock; --no-lock)");
    }

    Console.WriteLine();

    // --- 9. Verify (read-back) ---
    Console.WriteLine("Verification (read-back):");
    ctx.Load(list.ContentTypes, cts => cts.Include(c => c.Name, c => c.StringId));
    ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName, f => f.TypeAsString));
    await ctx.ExecuteQueryRetryAsync();

    var verifyCt = list.ContentTypes.FirstOrDefault(c =>
        string.Equals(c.Name, spOptions.DocumentSetContentType, StringComparison.OrdinalIgnoreCase));
    var ctOk = verifyCt is not null && verifyCt.StringId.StartsWith(DocSetBaseId, StringComparison.OrdinalIgnoreCase);
    Console.WriteLine($"  {(ctOk ? "✔" : (apply ? "✖" : "·"))} content type '{spOptions.DocumentSetContentType}' present as Document Set");
    foreach (var c in baseColumns)
    {
        var present = list.Fields.Any(f => f.InternalName == c);
        Console.WriteLine($"  {(present ? "✔" : (apply ? "✖" : "·"))} column '{c}'");
    }
    Console.WriteLine();
}

// -----------------------------------------------------------------------------
// 10. Emit the ready-to-paste config (PREPARED, LEFT DORMANT until go-live).
// -----------------------------------------------------------------------------
PrintConfigBlock(spOptions, opts.Practice, opts.Leader, opts.ToSite, library, parentFolder);

if (!apply)
{
    Console.WriteLine();
    Console.WriteLine("DRY RUN — nothing was written. Re-run with --apply to provision.");
}

return 0;

// =============================================================================
// Helpers
// =============================================================================

static string Tag(bool apply) => apply ? "  →" : "  would";

// Ensures the "Document Sets" site-collection feature is active, so the built-in Document Set content
// type (0x0120D520) is present in the web's content-type gallery to derive from. Idempotent.
async Task EnsureDocumentSetFeatureAsync(ClientContext ctx)
{
    var docSetFeatureId = new Guid("3bae86a2-776d-499d-9db8-fa4cdc7884f8");
    var features = ctx.Site.Features;
    ctx.Load(features);
    await ctx.ExecuteQueryRetryAsync();

    if (features.Any(f => f.DefinitionId == docSetFeatureId))
    {
        return;
    }

    Console.WriteLine("  activating the 'Document Sets' site-collection feature…");
    features.Add(docSetFeatureId, force: true, featdefScope: FeatureDefinitionScope.None);
    await ctx.ExecuteQueryRetryAsync();
}

async Task<SourceModel> ReadSourceAsync(
    SharePointContextFactory factory, string siteUrl, string library, string ctName, string[] wantedColumns)
{
    using var ctx = await factory.CreateContextAsync(siteUrl);
    ctx.Load(ctx.Web, w => w.WebTemplate, w => w.Configuration);
    var list = ctx.Web.Lists.GetByTitle(library);
    ctx.Load(list, l => l.EnableVersioning, l => l.MajorVersionLimit);
    ctx.Load(list.ContentTypes, cts => cts.Include(c => c.Name, c => c.StringId, c => c.Parent.StringId));
    ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName, f => f.TypeAsString, f => f.SchemaXml));
    await ctx.ExecuteQueryRetryAsync();

    var ct = list.ContentTypes.FirstOrDefault(c => string.Equals(c.Name, ctName, StringComparison.OrdinalIgnoreCase))
        ?? throw new InvalidOperationException($"Content type '{ctName}' not found on source library '{library}'.");

    var fields = new List<SourceField>();
    foreach (var col in wantedColumns)
    {
        var field = list.Fields.FirstOrDefault(f => f.InternalName == col);
        if (field is not null)
        {
            fields.Add(new SourceField(field.InternalName, field.TypeAsString, SanitizeFieldSchemaXml(field.SchemaXml)));
        }
    }

    return new SourceModel(
        WebTemplate: $"{ctx.Web.WebTemplate}#{ctx.Web.Configuration}",
        ParentContentTypeId: ct.Parent.StringId,
        EnableVersioning: list.EnableVersioning,
        MajorVersionLimit: list.MajorVersionLimit,
        Fields: fields);
}

// Strips site-specific attributes so the field can be re-created cleanly on another site while KEEPING the
// same internal name / id (so it stays a true mirror the global config already points at).
static string SanitizeFieldSchemaXml(string schemaXml)
{
    var el = XElement.Parse(schemaXml);
    foreach (var attr in new[] { "Version", "SourceID", "WebId", "List" })
    {
        // Keep List only for lookup/user fields that genuinely need it (UserInfo); drop stale cross-site refs otherwise.
        if (attr == "List" && (string?)el.Attribute("Type") is "User" or "UserMulti")
        {
            continue;
        }

        el.Attribute(attr)?.Remove();
    }

    return el.ToString(SaveOptions.DisableFormatting);
}

async Task LockFolderAsync(ClientContext ctx, string currentUrl)
{
    var folder = ctx.Web.GetFolderByServerRelativeUrl(currentUrl);
    var item = folder.ListItemAllFields;
    ctx.Load(item, i => i.HasUniqueRoleAssignments);
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
        a => a.Member.PrincipalType, a => a.Member.Id, a => a.Member.Title,
        a => a.RoleDefinitionBindings.Include(rd => rd.BasePermissions)));
    await ctx.ExecuteQueryRetryAsync();

    var ownerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
    var toFix = item.RoleAssignments.Where(a =>
            a.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup &&
            a.Member.Id != ownerGroupId &&
            a.RoleDefinitionBindings.Any(rd => rd.BasePermissions.Has(PermissionKind.AddListItems)))
        .ToList();

    foreach (var ra in toFix)
    {
        ra.RoleDefinitionBindings.RemoveAll();
        ra.RoleDefinitionBindings.Add(readDef);
        ra.Update();
    }
    await ctx.ExecuteQueryRetryAsync();
    Console.WriteLine($"✔ locked '{currentUrl}': {toFix.Count} group(s) downgraded to Read (only the app + Owners can create).");
}

static void PrintBootstrapHelp(string toSite, string title, string fromSite, string clientId, string tenant)
{
    var uri = new Uri(toSite);
    Console.Error.WriteLine("This tool runs app-only with the sync certificate, which can only reach sites it has been");
    Console.Error.WriteLine("granted on. Two one-time, admin-only steps come first (no app registration needed):");
    Console.Error.WriteLine();
    Console.Error.WriteLine($"  1. Create the empty site \"{title}\" at {toSite}");
    Console.Error.WriteLine($"  2. Grant app {clientId} a Sites.Selected fullControl grant on it");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Run (Windows PowerShell 5.1):");
    Console.Error.WriteLine("  powershell -ExecutionPolicy Bypass -File .\\tools\\ProvisionPracticeSite\\bootstrap-practice-site.ps1 `");
    Console.Error.WriteLine($"      -SiteUrl {toSite} -Title \"{title}\" -OwnerUpn <your-admin-upn>");
    Console.Error.WriteLine("Then re-run this tool with --apply.");
}

static void PrintConfigBlock(SharePointOptions sp, string practice, string? leader, string toSite, string library, string parentFolder)
{
    // Next free PracticeMappings index (existing config has :0 for Estate & Gift).
    var idx = Math.Max(1, sp.PracticeMappings.Count);
    Console.WriteLine("─────────────────────────────────────────────────────────────────────────────");
    Console.WriteLine("Config to add (PREPARED — leave DORMANT until the leaders are briefed):");
    Console.WriteLine();
    Console.WriteLine("  local.settings.json (\"Values\"):");
    Console.WriteLine($"    \"SharePoint:PracticeMappings:{idx}:Practice\": \"{practice}\",");
    if (!string.IsNullOrWhiteSpace(leader))
        Console.WriteLine($"    \"SharePoint:PracticeMappings:{idx}:PracticeLeaderEmail\": \"{leader}\",");
    Console.WriteLine($"    \"SharePoint:PracticeMappings:{idx}:SiteUrl\": \"{toSite}\",");
    Console.WriteLine($"    \"SharePoint:PracticeMappings:{idx}:Library\": \"{library}\",");
    Console.WriteLine($"    \"SharePoint:PracticeMappings:{idx}:ParentFolder\": \"{parentFolder}\",");
    Console.WriteLine("    (AdminEmails intentionally omitted — no practice admin yet.)");
    Console.WriteLine();
    Console.WriteLine("  Azure app settings (env-var form): same keys with '__' separators, e.g.");
    Console.WriteLine($"    SharePoint__PracticeMappings__{idx}__Practice = {practice}");
    Console.WriteLine();
    Console.WriteLine("  GO-LIVE SWITCH (add ONLY when ready to start syncing this practice):");
    Console.WriteLine($"    Acumatica:IncludedPractices:1 = {practice}");
    Console.WriteLine($"    HubSpot:IncludedPractices:1    = {practice}");
    Console.WriteLine("  Until those two allow-list entries exist, the sync ignores this practice entirely —");
    Console.WriteLine("  the site simply sits ready.");
    Console.WriteLine("─────────────────────────────────────────────────────────────────────────────");
}

static ProvArgs? ParseArgs(string[] args)
{
    string? practice = null, leader = null, to = null, from = null, library = null, parent = null, title = null;
    var apply = false;
    var lockFolder = true;

    for (var i = 0; i < args.Length; i++)
    {
        switch (args[i].ToLowerInvariant())
        {
            case "--practice": practice = Next(args, ref i); break;
            case "--leader": leader = Next(args, ref i); break;
            case "--to": to = Next(args, ref i); break;
            case "--from": from = Next(args, ref i); break;
            case "--library": library = Next(args, ref i); break;
            case "--parent-folder": parent = Next(args, ref i); break;
            case "--title": title = Next(args, ref i); break;
            case "--apply": apply = true; break;
            case "--no-lock": lockFolder = false; break;
            default:
                Console.Error.WriteLine($"Unknown argument: {args[i]}");
                return null;
        }
    }

    if (string.IsNullOrWhiteSpace(practice) || string.IsNullOrWhiteSpace(to))
    {
        Console.Error.WriteLine("Usage: dotnet run --project tools/ProvisionPracticeSite -- \\");
        Console.Error.WriteLine("         --practice \"Marital Dissolution\" --leader rhoffman@marshall-stevens.com \\");
        Console.Error.WriteLine("         --to https://marshallstevens.sharepoint.com/sites/MaritalDissolution \\");
        Console.Error.WriteLine("         [--from <source site>] [--library Documents] [--parent-folder Projects/Current] \\");
        Console.Error.WriteLine("         [--no-lock] [--apply]");
        return null;
    }

    return new ProvArgs(practice!.Trim(), leader?.Trim(), to!.Trim(), from?.Trim(), library?.Trim(), parent?.Trim(), title?.Trim(), apply, lockFolder);

    static string? Next(string[] a, ref int i) => (i + 1 < a.Length) ? a[++i] : null;
}

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

sealed record ProvArgs(
    string Practice, string? Leader, string ToSite, string? FromSite,
    string? Library, string? ParentFolder, string? Title, bool Apply, bool Lock);

sealed record SourceField(string InternalName, string Type, string SchemaXml);

sealed record SourceModel(
    string WebTemplate, string ParentContentTypeId, bool EnableVersioning, int MajorVersionLimit,
    List<SourceField> Fields);
