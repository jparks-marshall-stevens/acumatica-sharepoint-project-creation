using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using SPDocumentSet = Microsoft.SharePoint.Client.DocumentSet.DocumentSet;

namespace ProjectSync.SharePoint;

public sealed class SharePointDocumentSetService : ISharePointDocumentSetService
{
    private readonly SharePointContextFactory _contextFactory;
    private readonly SharePointOptions _options;
    private readonly ILogger<SharePointDocumentSetService> _logger;

    public SharePointDocumentSetService(
        SharePointContextFactory contextFactory,
        IOptions<SharePointOptions> options,
        ILogger<SharePointDocumentSetService> logger)
    {
        _contextFactory = contextFactory;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<DocumentSetResult> EnsureProjectDocumentSetAsync(
        AcumaticaProject project,
        CancellationToken cancellationToken)
    {
        var destination = ResolveDestination(project.Practice);
        var siteUrl = string.IsNullOrWhiteSpace(destination.SiteUrl) ? _options.SiteUrl : destination.SiteUrl!;

        using var ctx = await _contextFactory.CreateContextAsync(siteUrl);

        var list = ctx.Web.Lists.GetByTitle(destination.Library);
        ctx.Load(list, l => l.RootFolder.ServerRelativeUrl, l => l.ContentTypes);
        await ctx.ExecuteQueryRetryAsync();

        var setName = SharePointNaming.SanitizeLeafName(project.ProjectId);

        // Idempotency: bail out if a folder/document set with this name already exists in the library.
        var existing = await FindExistingAsync(ctx, list, setName, cancellationToken);
        if (existing is not null)
        {
            _logger.LogInformation("Document set '{Name}' already exists at {Url}; updating metadata only.",
                setName, existing);
            await ApplyMetadataAsync(ctx, existing, project, cancellationToken);
            return new DocumentSetResult(Created: false, existing);
        }

        var parentFolder = list.RootFolder;
        if (!string.IsNullOrWhiteSpace(destination.ParentFolder))
        {
            parentFolder = EnsureFolderPath(ctx, list, destination.ParentFolder!);
        }

        var contentType = list.ContentTypes.FirstOrDefault(c =>
            string.Equals(c.Name, _options.DocumentSetContentType, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                $"Content type '{_options.DocumentSetContentType}' is not enabled on library '{destination.Library}'.");

        _logger.LogInformation("Creating document set '{Name}' in {Library} (practice '{Practice}').",
            setName, destination.Library, project.Practice ?? "<none>");

        var created = SPDocumentSet.Create(ctx, parentFolder, setName, contentType.Id);
        await ctx.ExecuteQueryRetryAsync();

        var serverRelativeUrl = created.Value;
        await ApplyMetadataAsync(ctx, serverRelativeUrl, project, cancellationToken);

        return new DocumentSetResult(Created: true, serverRelativeUrl);
    }

    private PracticeDestination ResolveDestination(string? practice)
    {
        var key = practice?.Trim();
        if (!string.IsNullOrEmpty(key) && _options.PracticeMappings.TryGetValue(key, out var dest))
        {
            return dest;
        }

        if (_options.PracticeMappings.TryGetValue("*", out var fallback))
        {
            _logger.LogWarning("No practice mapping for '{Practice}'; using default '*' mapping.", practice ?? "<none>");
            return fallback;
        }

        throw new InvalidOperationException(
            $"No SharePoint destination mapped for practice '{practice ?? "<none>"}' and no '*' default was configured.");
    }

    private static async Task<string?> FindExistingAsync(
        ClientContext ctx, List list, string leafName, CancellationToken cancellationToken)
    {
        var safeValue = leafName.Replace("'", "&apos;");
        var query = new CamlQuery
        {
            ViewXml =
                "<View Scope='RecursiveAll'><Query><Where><And>" +
                $"<Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>{safeValue}</Value></Eq>" +
                "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                "</And></Where></Query><RowLimit>1</RowLimit></View>",
        };

        var items = list.GetItems(query);
        ctx.Load(items, c => c.Include(i => i.FileSystemObjectType, i => i["FileRef"]));
        await ctx.ExecuteQueryRetryAsync();

        return items.Count > 0 ? items[0]["FileRef"]?.ToString() : null;
    }

    private static Folder EnsureFolderPath(ClientContext ctx, List list, string relativePath)
    {
        var current = list.RootFolder;
        foreach (var segment in relativePath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var name = SharePointNaming.SanitizeLeafName(segment);
            current = current.Folders.Add(name);
            ctx.Load(current);
        }

        ctx.ExecuteQueryRetry();
        return current;
    }

    private async Task ApplyMetadataAsync(
        ClientContext ctx, string serverRelativeUrl, AcumaticaProject project, CancellationToken cancellationToken)
    {
        var folder = ctx.Web.GetFolderByServerRelativeUrl(serverRelativeUrl);
        var item = folder.ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();

        SetIfPresent(item, _options.ProjectIdColumn, project.ProjectId);
        SetIfPresent(item, _options.CustomerNameColumn, project.CustomerName);
        SetIfPresent(item, _options.ProjectNameColumn, project.ProjectName);
        SetIfPresent(item, _options.ProjectManagerColumn, project.ProjectManager);
        SetIfPresent(item, _options.PracticeColumn, project.Practice);

        item.Update();
        await ctx.ExecuteQueryRetryAsync();
    }

    private void SetIfPresent(ListItem item, string column, string? value)
    {
        if (string.IsNullOrWhiteSpace(column) || value is null)
        {
            return;
        }

        if (!item.FieldValues.ContainsKey(column))
        {
            _logger.LogWarning("Column '{Column}' not found on the document set list item; skipping.", column);
            return;
        }

        item[column] = value;
    }
}
