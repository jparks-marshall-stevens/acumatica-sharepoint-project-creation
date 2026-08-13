using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;

// -----------------------------------------------------------------------------
// SharePoint connectivity test — READ ONLY. Creates nothing.
//
// Authenticates app-only (certificate) to the configured site and reports:
//   • the web it reached (proves auth + Sites.Selected grant)
//   • each document library, whether the Document Set content type is enabled,
//     and the custom columns with their INTERNAL names + types
//   • a check of the configured metadata column internal names
//
// Config comes from the Functions local.settings.json "Values" section.
// -----------------------------------------------------------------------------

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var options = new SharePointOptions();
configuration.GetSection(SharePointOptions.SectionName).Bind(options);

using var loggerFactory = LoggerFactory.Create(b => b.SetMinimumLevel(LogLevel.Warning).AddSimpleConsole(o => o.SingleLine = true));
var wrapper = Microsoft.Extensions.Options.Options.Create(options);
var factory = new SharePointContextFactory(wrapper, loggerFactory.CreateLogger<SharePointContextFactory>());

Console.WriteLine("=== SharePoint connectivity test (READ ONLY) ===");
Console.WriteLine($"Site      : {options.SiteUrl}");
Console.WriteLine($"Client ID : {options.ClientId}");
Console.WriteLine($"Cert      : thumbprint {options.CertificateThumbprint}");
Console.WriteLine($"Doc set CT: {options.DocumentSetContentType}");
Console.WriteLine();

try
{
    using var ctx = await factory.CreateContextAsync(options.SiteUrl);

    ctx.Load(ctx.Web, w => w.Title, w => w.Url);
    await ctx.ExecuteQueryRetryAsync();
    Console.WriteLine($"✔ Connected. Web: \"{ctx.Web.Title}\"  ({ctx.Web.Url})");
    Console.WriteLine();

    var lists = ctx.Web.Lists;
    ctx.Load(lists, ls => ls.Include(
        l => l.Title, l => l.BaseTemplate, l => l.Hidden, l => l.ItemCount, l => l.ContentTypesEnabled));
    await ctx.ExecuteQueryRetryAsync();

    // Document libraries only (template 101), excluding hidden system libraries.
    var docLibs = lists.Where(l => l.BaseTemplate == 101 && !l.Hidden).ToList();

    foreach (var list in docLibs)
    {
        ctx.Load(list.ContentTypes, c => c.Include(ct => ct.Name, ct => ct.StringId));
        ctx.Load(list.Fields, f => f.Include(
            x => x.InternalName, x => x.Title, x => x.TypeAsString, x => x.Hidden, x => x.FromBaseType));
    }
    await ctx.ExecuteQueryRetryAsync();

    // Content type IDs that start with this are Document Set types (OOTB or custom-derived).
    const string docSetBaseId = "0x0120D520";

    var configured = new[]
    {
        ("ProjectIdColumn", options.ProjectIdColumn),
        ("CustomerNameColumn", options.CustomerNameColumn),
        ("ProjectNameColumn", options.ProjectNameColumn),
        ("ProjectManagerColumn", options.ProjectManagerColumn),
        ("PracticeColumn", options.PracticeColumn),
    };

    Console.WriteLine($"Document libraries ({docLibs.Count}):");
    Console.WriteLine();
    foreach (var list in docLibs)
    {
        Console.WriteLine($"  ▸ \"{list.Title}\"  (items: {list.ItemCount}, content types: {(list.ContentTypesEnabled ? "on" : "off")})");

        Console.WriteLine("      Content types (name — docset?):");
        foreach (var ct in list.ContentTypes)
        {
            var isDocSet = ct.StringId.StartsWith(docSetBaseId, StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"        {ct.Name,-24} {(isDocSet ? "[Document Set]" : "")}");
        }

        var configuredCt = list.ContentTypes.FirstOrDefault(c =>
            string.Equals(c.Name, options.DocumentSetContentType, StringComparison.OrdinalIgnoreCase));
        var ctOk = configuredCt is not null;
        var ctIsDocSet = configuredCt is not null && configuredCt.StringId.StartsWith(docSetBaseId, StringComparison.OrdinalIgnoreCase);
        Console.WriteLine($"      Configured CT '{options.DocumentSetContentType}': {(ctOk ? (ctIsDocSet ? "✔ present, is a Document Set" : "⚠ present but NOT a Document Set type") : "✖ not on this library")}");

        // Custom columns only (not inherited base fields, not hidden).
        var custom = list.Fields
            .Where(f => !f.FromBaseType && !f.Hidden)
            .OrderBy(f => f.InternalName)
            .ToList();
        if (custom.Count > 0)
        {
            Console.WriteLine("      Custom columns (internal name : type — \"title\"):");
            foreach (var f in custom)
            {
                Console.WriteLine($"        {f.InternalName,-24} : {f.TypeAsString,-12} \"{f.Title}\"");
            }
        }

        // Check the configured metadata column internal names against this library.
        var present = configured
            .Select(c => (c.Item1, c.Item2, Found: list.Fields.Any(f => f.InternalName == c.Item2)))
            .ToList();
        if (present.Any(p => p.Found))
        {
            Console.WriteLine("      Configured columns match here:");
            foreach (var (label, internalName, found) in present)
            {
                Console.WriteLine($"        {(found ? "✔" : "✖")} {label,-22} = '{internalName}'");
            }
        }

        Console.WriteLine();
    }

    Console.WriteLine("✅ SharePoint connectivity OK (no changes made).");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ SharePoint connectivity failed: {ex.Message}");
    if (ex.InnerException is not null)
    {
        Console.Error.WriteLine($"   inner: {ex.InnerException.Message}");
    }
    return 1;
}

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!System.IO.File.Exists(path))
    {
        return result;
    }

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
