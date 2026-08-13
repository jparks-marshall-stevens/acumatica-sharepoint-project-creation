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

        // Idempotency is keyed on the Project Id metadata column (unique), NOT the folder name —
        // folder names come from the description and are not guaranteed unique.
        var existing = await FindExistingByProjectIdAsync(ctx, list, project.ProjectId, cancellationToken);
        if (existing is not null)
        {
            _logger.LogInformation("Document set for project {ProjectId} already exists at {Url}; updating metadata + permissions.",
                project.ProjectId, existing);
            await ApplyMetadataAsync(ctx, existing, project, cancellationToken);
            await ApplyPermissionsAsync(ctx, existing, project, destination, cancellationToken);
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

        // Name = "{first N chars of customer name} | {project id}" (unique because the id is included).
        var desiredName = SharePointNaming.BuildDocumentSetName(
            project.CustomerName, project.ProjectId, _options.DocumentSetNameMaxLength);
        var setName = await ResolveUniqueNameAsync(ctx, parentFolder, desiredName, project.ProjectId, cancellationToken);

        _logger.LogInformation("Creating document set '{Name}' for project {ProjectId} in {Library} (practice '{Practice}').",
            setName, project.ProjectId, destination.Library, project.Practice ?? "<none>");

        var created = SPDocumentSet.Create(ctx, parentFolder, setName, contentType.Id);
        await ctx.ExecuteQueryRetryAsync();

        var serverRelativeUrl = created.Value;
        await ApplyMetadataAsync(ctx, serverRelativeUrl, project, cancellationToken);
        await ApplyPermissionsAsync(ctx, serverRelativeUrl, project, destination, cancellationToken);

        return new DocumentSetResult(Created: true, serverRelativeUrl);
    }

    public DocumentSetPlan PlanDocumentSet(AcumaticaProject project)
    {
        var destination = ResolveDestination(project.Practice);
        var siteUrl = string.IsNullOrWhiteSpace(destination.SiteUrl) ? _options.SiteUrl : destination.SiteUrl!;
        return new DocumentSetPlan(
            siteUrl,
            destination.Library,
            destination.ParentFolder,
            SharePointNaming.BuildDocumentSetName(project.CustomerName, project.ProjectId, _options.DocumentSetNameMaxLength));
    }

    private PracticeMappingEntry ResolveDestination(string? practice)
    {
        var key = practice?.Trim();
        if (!string.IsNullOrEmpty(key))
        {
            var match = _options.PracticeMappings
                .FirstOrDefault(m => string.Equals(m.Practice.Trim(), key, StringComparison.OrdinalIgnoreCase));
            if (match is not null)
            {
                return match;
            }
        }

        var fallback = _options.PracticeMappings.FirstOrDefault(m => m.Practice == "*");
        if (fallback is not null)
        {
            _logger.LogWarning("No practice mapping for '{Practice}'; using default '*' mapping.", practice ?? "<none>");
            return fallback;
        }

        throw new InvalidOperationException(
            $"No SharePoint destination mapped for practice '{practice ?? "<none>"}' and no '*' default was configured.");
    }

    private async Task<string?> FindExistingByProjectIdAsync(
        ClientContext ctx, List list, string projectId, CancellationToken cancellationToken)
    {
        var safeValue = System.Security.SecurityElement.Escape(projectId) ?? projectId;
        var query = new CamlQuery
        {
            ViewXml =
                "<View Scope='RecursiveAll'><Query><Where><And>" +
                $"<Eq><FieldRef Name='{_options.ProjectIdColumn}'/><Value Type='Text'>{safeValue}</Value></Eq>" +
                "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
                "</And></Where></Query><RowLimit>1</RowLimit></View>",
        };

        var items = list.GetItems(query);
        ctx.Load(items, c => c.Include(i => i.FileSystemObjectType, i => i["FileRef"]));
        await ctx.ExecuteQueryRetryAsync();

        return items.Count > 0 ? items[0]["FileRef"]?.ToString() : null;
    }

    /// <summary>
    /// Returns a folder name unique among <paramref name="parentFolder"/>'s child folders. If the
    /// desired name is taken (different project, same description prefix), appends the project id.
    /// </summary>
    private static async Task<string> ResolveUniqueNameAsync(
        ClientContext ctx, Folder parentFolder, string desiredName, string projectId, CancellationToken cancellationToken)
    {
        ctx.Load(parentFolder.Folders, fs => fs.Include(f => f.Name));
        await ctx.ExecuteQueryRetryAsync();

        var taken = new HashSet<string>(
            parentFolder.Folders.Select(f => f.Name), StringComparer.OrdinalIgnoreCase);

        if (!taken.Contains(desiredName))
        {
            return desiredName;
        }

        var withId = SharePointNaming.SanitizeLeafName($"{desiredName} - {projectId}");
        if (!taken.Contains(withId))
        {
            return withId;
        }

        for (var i = 2; ; i++)
        {
            var candidate = SharePointNaming.SanitizeLeafName($"{desiredName} - {projectId} ({i})");
            if (!taken.Contains(candidate))
            {
                return candidate;
            }
        }
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
        var folder = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl));
        var item = folder.ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();

        // Resolve the PM person FIRST — it needs its own round-trip. Doing this before setting any
        // fields keeps all field writes in a single Update()/ExecuteQuery, so no staged text value
        // is lost to a mid-sequence round-trip.
        FieldUserValue? pmValue = null;
        if (_options.ProjectManagerIsPersonColumn)
        {
            pmValue = await ResolvePersonAsync(ctx, _options.ProjectManagerColumn, item,
                project.ProjectManagerEmail, project.ProjectManager);
        }

        SetIfPresent(item, _options.ProjectIdColumn, project.ProjectId);
        SetIfPresent(item, _options.CustomerNameColumn, project.CustomerName);
        SetIfPresent(item, _options.ProjectNameColumn, project.ProjectName);
        SetIfPresent(item, _options.PracticeColumn, project.Practice);

        if (_options.ProjectManagerIsPersonColumn)
        {
            if (pmValue is not null && item.FieldValues.ContainsKey(_options.ProjectManagerColumn))
            {
                item[_options.ProjectManagerColumn] = pmValue;
            }
        }
        else
        {
            SetIfPresent(item, _options.ProjectManagerColumn, project.ProjectManager);
        }

        item.Update();
        await ctx.ExecuteQueryRetryAsync();
    }

    /// <summary>
    /// Resolves a person by email/UPN (preferred) or name to a <see cref="FieldUserValue"/>.
    /// Returns null (fail-soft) if the column is absent or the identity can't be resolved.
    /// </summary>
    private async Task<FieldUserValue?> ResolvePersonAsync(
        ClientContext ctx, string column, ListItem item, string? email, string? displayName)
    {
        if (string.IsNullOrWhiteSpace(column) || !item.FieldValues.ContainsKey(column))
        {
            _logger.LogWarning("People column '{Column}' not found on the document set list item; skipping.", column);
            return null;
        }

        // Prefer the email/UPN — display names don't resolve reliably (and can be ambiguous).
        var identity = !string.IsNullOrWhiteSpace(email) ? email : displayName;
        if (string.IsNullOrWhiteSpace(email) && !string.IsNullOrWhiteSpace(displayName))
        {
            _logger.LogWarning(
                "No PM email available; attempting to resolve people field '{Column}' by display name '{Name}', which may fail.",
                column, displayName);
        }

        var user = await TryEnsureUserAsync(ctx, identity);
        return user is null ? null : new FieldUserValue { LookupId = user.Id };
    }

    /// <summary>Resolves an email/UPN to a site user via EnsureUser. Returns null (fail-soft) if unresolvable.</summary>
    private async Task<User?> TryEnsureUserAsync(ClientContext ctx, string? identity)
    {
        if (string.IsNullOrWhiteSpace(identity))
        {
            return null;
        }

        try
        {
            var user = ctx.Web.EnsureUser(identity);
            ctx.Load(user, u => u.Id);
            await ctx.ExecuteQueryRetryAsync();
            return user;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not resolve '{Identity}' to a SharePoint user.", identity);
            return null;
        }
    }

    /// <summary>
    /// Sets access on the document set: grants the project manager and practice leader the configured
    /// level; when RestrictPermissions is on, breaks inheritance so only they and the site Owners have
    /// access. Fail-soft on individual unresolvable users; a hard permission-API failure is logged as an
    /// error (surfaces via alerting) but does not roll back the created set.
    /// </summary>
    private async Task ApplyPermissionsAsync(
        ClientContext ctx, string serverRelativeUrl, AcumaticaProject project, PracticeMappingEntry destination, CancellationToken cancellationToken)
    {
        if (!_options.SetProjectPermissions)
        {
            return;
        }

        try
        {
            var web = ctx.Web;
            var folder = web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl));
            var item = folder.ListItemAllFields;
            var grantRole = web.RoleDefinitions.GetByName(_options.PermissionLevel);
            var fullControl = web.RoleDefinitions.GetByType(RoleType.Administrator);
            ctx.Load(item);
            ctx.Load(grantRole);
            ctx.Load(fullControl);
            ctx.Load(web.AssociatedOwnerGroup);
            await ctx.ExecuteQueryRetryAsync();

            var pmIdentity = !string.IsNullOrWhiteSpace(project.ProjectManagerEmail)
                ? project.ProjectManagerEmail
                : project.ProjectManager;
            var pmUser = await TryEnsureUserAsync(ctx, pmIdentity);
            var leaderUser = await TryEnsureUserAsync(ctx, destination.PracticeLeaderEmail);

            if (_options.RestrictPermissions)
            {
                // Clean slate, then re-grant. Site collection admins / the app identity retain access.
                item.BreakRoleInheritance(copyRoleAssignments: false, clearSubscopes: true);
                var ownerBinding = new RoleDefinitionBindingCollection(ctx);
                ownerBinding.Add(fullControl);
                item.RoleAssignments.Add(web.AssociatedOwnerGroup, ownerBinding);
            }

            foreach (var user in new[] { pmUser, leaderUser })
            {
                if (user is null)
                {
                    continue;
                }

                var binding = new RoleDefinitionBindingCollection(ctx);
                binding.Add(grantRole);
                item.RoleAssignments.Add(user, binding);
            }

            await ctx.ExecuteQueryRetryAsync();
            _logger.LogInformation(
                "Set permissions on document set for project {ProjectId} (restrict={Restrict}, pm={Pm}, leader={Leader}).",
                project.ProjectId, _options.RestrictPermissions, pmUser is not null, leaderUser is not null);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "Failed to set permissions on the document set for project {ProjectId}. The set exists with its current permissions; re-run to retry.",
                project.ProjectId);
        }
    }

    /// <summary>DocumentSet.Create can return an absolute URL; CSOM folder lookup needs the decoded server-relative path.</summary>
    private static string ToServerRelative(string url)
        => url.StartsWith("http", StringComparison.OrdinalIgnoreCase)
            ? Uri.UnescapeDataString(new Uri(url).AbsolutePath)
            : url;

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
