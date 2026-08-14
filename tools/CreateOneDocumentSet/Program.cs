using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// -----------------------------------------------------------------------------
// Create ONE document set — a controlled live write for a single project id.
//
// Runs the exact production path (SharePointDocumentSetService.EnsureProjectDocumentSetAsync),
// then reads the metadata back to verify what landed (including the PM People field).
//
// Usage: dotnet run --project tools/CreateOneDocumentSet -- <ProjectId>
// -----------------------------------------------------------------------------

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: dotnet run --project tools/CreateOneDocumentSet -- <ProjectId>");
    return 1;
}

var projectId = args[0].Trim();

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var acumaticaOptions = Bind<AcumaticaOptions>(configuration, AcumaticaOptions.SectionName);
var sharePointOptions = Bind<SharePointOptions>(configuration, SharePointOptions.SectionName);

using var loggerFactory = LoggerFactory.Create(b => b
    .SetMinimumLevel(LogLevel.Warning)
    .AddFilter("ProjectSync", LogLevel.Information) // surface our own info logs (e.g. upload-link creation)
    .AddSimpleConsole(o => o.SingleLine = true));
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(acumaticaOptions.Value.TimeoutSeconds) };

var tokenProvider = new AcumaticaTokenProvider(http, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaTokenProvider>());
var acumatica = new AcumaticaClient(http, tokenProvider, acumaticaOptions, loggerFactory.CreateLogger<AcumaticaClient>());
var contextFactory = new SharePointContextFactory(sharePointOptions, loggerFactory.CreateLogger<SharePointContextFactory>());
var uploadLinks = new GraphUploadLinkService(contextFactory, sharePointOptions, loggerFactory.CreateLogger<GraphUploadLinkService>());
var sharePoint = new SharePointDocumentSetService(contextFactory, uploadLinks, sharePointOptions, loggerFactory.CreateLogger<SharePointDocumentSetService>());

Console.WriteLine($"=== Create ONE document set: project {projectId} ===");
Console.WriteLine();

// 1. Find the project in the GI (wide window).
var projects = await acumatica.GetProjectsCreatedAfterAsync(DateTimeOffset.UtcNow.AddYears(-5), CancellationToken.None);
var project = projects.FirstOrDefault(p => string.Equals(p.ProjectId, projectId, StringComparison.OrdinalIgnoreCase));
if (project is null)
{
    Console.Error.WriteLine($"❌ Project '{projectId}' not found in the GI.");
    return 2;
}

// Enrich with the project team (emails granted alongside PM + leader).
var teamEmails = await acumatica.GetTeamEmailsAsync(project.ProjectId, CancellationToken.None);
project = project with { TeamEmails = teamEmails };

var plan = sharePoint.PlanDocumentSet(project);
Console.WriteLine("Project:");
Console.WriteLine($"    Team ({teamEmails.Count})   : {string.Join(", ", teamEmails)}");
Console.WriteLine($"    Project Id      : {project.ProjectId}");
Console.WriteLine($"    Project Name    : {project.ProjectName}");
Console.WriteLine($"    Customer Name   : {project.CustomerName}");
Console.WriteLine($"    Project Manager : {project.ProjectManager}  (email: {project.ProjectManagerEmail})");
Console.WriteLine($"    Practice        : {project.Practice}");
Console.WriteLine($"    → Would create  : \"{plan.SetName}\" in {plan.Library} @ {plan.SiteUrl}");
Console.WriteLine();

// Optional: recycle the existing document set for this project (soft delete → site Recycle Bin).
if (args.Contains("--delete"))
{
    using var delCtx = await contextFactory.CreateContextAsync(plan.SiteUrl);
    var delList = delCtx.Web.Lists.GetByTitle(plan.Library);
    var pidCol = sharePointOptions.Value.ProjectIdColumn;
    var query = new CamlQuery
    {
        ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>" +
                  $"<Eq><FieldRef Name='{pidCol}'/><Value Type='Text'>{project.ProjectId}</Value></Eq>" +
                  "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                  "</And></Where></Query><RowLimit>5</RowLimit></View>",
    };
    var items = delList.GetItems(query);
    delCtx.Load(items, i => i.Include(x => x["FileRef"]));
    await delCtx.ExecuteQueryRetryAsync();
    if (items.Count == 0)
    {
        Console.WriteLine("No existing document set found for this project id; nothing to delete.");
    }
    foreach (var it in items)
    {
        var url = it["FileRef"]?.ToString();
        Console.WriteLine($"Recycling: {url}");
        delCtx.Web.GetFolderByServerRelativeUrl(url).Recycle();
    }
    await delCtx.ExecuteQueryRetryAsync();
    Console.WriteLine("✔ Recycled to the site Recycle Bin (recoverable).");
    return 0;
}

// 2. Create (or update if it already exists).
Console.WriteLine("Creating document set…");
var result = await sharePoint.EnsureProjectDocumentSetAsync(project, CancellationToken.None);
Console.WriteLine($"✔ {(result.Created ? "CREATED" : "already existed — metadata updated")}: {result.ServerRelativeUrl}");
Console.WriteLine();

// 3. Read the metadata back to verify what actually landed.
Console.WriteLine("Verifying written metadata…");
using var ctx = await contextFactory.CreateContextAsync(plan.SiteUrl);
var folder = ctx.Web.GetFolderByServerRelativeUrl(result.ServerRelativeUrl);
var item = folder.ListItemAllFields;
ctx.Load(item);
await ctx.ExecuteQueryRetryAsync();

string Show(string col)
{
    if (!item.FieldValues.ContainsKey(col) || item[col] is null) return "<blank>";
    return item[col] switch
    {
        FieldUserValue u => $"{u.LookupValue} (id {u.LookupId})",
        _ => item[col]!.ToString() ?? "<blank>",
    };
}

Console.WriteLine($"    {sharePointOptions.Value.ProjectIdColumn,-24} = {Show(sharePointOptions.Value.ProjectIdColumn)}");
Console.WriteLine($"    {sharePointOptions.Value.CustomerNameColumn,-24} = {Show(sharePointOptions.Value.CustomerNameColumn)}");
Console.WriteLine($"    {sharePointOptions.Value.ProjectNameColumn,-24} = {Show(sharePointOptions.Value.ProjectNameColumn)}");
Console.WriteLine($"    {sharePointOptions.Value.ProjectManagerColumn,-24} = {Show(sharePointOptions.Value.ProjectManagerColumn)}");

var pmResolved = item.FieldValues.TryGetValue(sharePointOptions.Value.ProjectManagerColumn, out var pmVal) && pmVal is FieldUserValue;
Console.WriteLine();
Console.WriteLine(pmResolved
    ? "✅ PM People field RESOLVED to a SharePoint user."
    : "⚠ PM People field is blank (email did not resolve to an in-tenant user).");

// Verify permissions on the document set.
Console.WriteLine();
Console.WriteLine("Verifying permissions…");
try
{
    ctx.Load(item, i => i.HasUniqueRoleAssignments);
    var ras = item.RoleAssignments;
    ctx.Load(ras, r => r.Include(
        a => a.Member.LoginName, a => a.Member.Title,
        a => a.RoleDefinitionBindings.Include(rd => rd.Name)));
    await ctx.ExecuteQueryRetryAsync();
    Console.WriteLine($"    Inheritance broken (unique perms): {item.HasUniqueRoleAssignments}");
    foreach (var ra in ras)
    {
        var roles = string.Join(", ", ra.RoleDefinitionBindings.Select(rd => rd.Name));
        Console.WriteLine($"    {ra.Member.Title,-32} : {roles}");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"    ⚠ Could not read/set permissions: {ex.Message}");
    Console.WriteLine("    (The app likely lacks Full Control on the site — upgrade the Sites.Selected grant to fullControl.)");
}

// Verify the Client Uploads folder + upload link when the feature is enabled.
if (sharePointOptions.Value.CreateClientUploadLink)
{
    Console.WriteLine();
    Console.WriteLine("Verifying Client Uploads folder + upload link…");
    var uploadsName = sharePointOptions.Value.ClientUploadsFolderName;
    ctx.Load(folder.Folders, fs => fs.Include(f => f.Name, f => f.ServerRelativeUrl));
    await ctx.ExecuteQueryRetryAsync();
    var uploads = folder.Folders.FirstOrDefault(f => string.Equals(f.Name, uploadsName, StringComparison.OrdinalIgnoreCase));
    Console.WriteLine(uploads is not null
        ? $"    ✔ Subfolder present : {uploads.ServerRelativeUrl}"
        : $"    ⚠ Subfolder '{uploadsName}' not found.");

    var linkCol = sharePointOptions.Value.ClientUploadLinkColumn;
    if (item.FieldValues.TryGetValue(linkCol, out var linkVal) && linkVal is not null)
    {
        var url = linkVal is FieldUrlValue fu ? fu.Url : linkVal.ToString();
        Console.WriteLine($"    ✔ Link column '{linkCol}' = {url}");
    }
    else
    {
        Console.WriteLine($"    ⚠ Link column '{linkCol}' is blank or absent (link may have been created but not stamped — add the column to see it).");
    }
}

Console.WriteLine();
Console.WriteLine($"Open in browser: {plan.SiteUrl}");
return 0;

static IOptions<T> Bind<T>(IConfiguration config, string section) where T : class, new()
{
    var value = new T();
    config.GetSection(section).Bind(value);
    return Microsoft.Extensions.Options.Options.Create(value);
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
