using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using ProjectSync.Acumatica;
using ProjectSync.Notifications;
using ProjectSync.Options;
using SPDocumentSet = Microsoft.SharePoint.Client.DocumentSet.DocumentSet;

namespace ProjectSync.SharePoint;

public sealed class SharePointDocumentSetService : ISharePointDocumentSetService
{
    private readonly SharePointContextFactory _contextFactory;
    private readonly GraphUploadLinkService _uploadLinks;
    private readonly WorkspaceNotifier _notifier;
    private readonly SharePointOptions _options;
    private readonly ILogger<SharePointDocumentSetService> _logger;

    public SharePointDocumentSetService(
        SharePointContextFactory contextFactory,
        GraphUploadLinkService uploadLinks,
        WorkspaceNotifier notifier,
        IOptions<SharePointOptions> options,
        ILogger<SharePointDocumentSetService> logger)
    {
        _contextFactory = contextFactory;
        _uploadLinks = uploadLinks;
        _notifier = notifier;
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

        await EnsureTextColumnAsync(ctx, list, _options.StatusColumn);
        await EnsureTextColumnAsync(ctx, list, _options.HubSpotDealIdColumn);
        await EnsureTextColumnAsync(ctx, list, _options.OpportunityIdColumn);
        await EnsureTextColumnAsync(ctx, list, _options.DataroomUrlColumn);

        // Idempotency is keyed on the Project Id metadata column (unique), NOT the folder name —
        // folder names come from the description and are not guaranteed unique.
        var existing = await FindExistingByProjectIdAsync(ctx, list, destination.ParentFolder, project.ProjectId, cancellationToken);
        if (existing is not null)
        {
            _logger.LogInformation("Document set for project {ProjectId} already exists at {Url}; updating metadata + permissions.",
                project.ProjectId, existing);
            var sig = ReconcileSignature.Compute(project, destination.PracticeLeaderEmail, destination.AdminEmails, destination.MembersPermissionLevel);
            await ApplyMetadataAsync(ctx, existing, project, sig, siteUrl, cancellationToken);
            var addedExisting = await ApplyPermissionsAsync(ctx, existing, project, destination, cancellationToken);
            await NotifyProjectAccessAddedAsync(ctx, project, destination, existing, siteUrl, addedExisting, cancellationToken);
            return new DocumentSetResult(Created: false, existing);
        }

        // Nothing carries this project id yet. If the conversion recorded a HubSpot identifier (the GI's
        // PQCode), the engagement probably already has a scoping workspace — created from the deal, before
        // the ERP knew about it — so PROMOTE that one in place rather than opening a second folder for the
        // same work. The value is matched against the opportunity number first, then the raw deal id, so it
        // links whichever of the two a person actually recorded.
        if (!string.IsNullOrWhiteSpace(project.HubSpotLink))
        {
            var scoping =
                await FindByColumnAsync(ctx, list, destination.ParentFolder, _options.OpportunityIdColumn, project.HubSpotLink!, cancellationToken)
                ?? await FindByColumnAsync(ctx, list, destination.ParentFolder, _options.HubSpotDealIdColumn, project.HubSpotLink!, cancellationToken);

            if (scoping is not null && !string.IsNullOrWhiteSpace(scoping.ProjectId))
            {
                // Two projects claiming one HubSpot identifier (almost certainly a typo). Don't hijack the
                // workspace that already belongs to the other project; fall through and create a fresh one.
                _logger.LogWarning(
                    "HubSpot link {Link} on project {ProjectId} points at a workspace already promoted to project {OwnerProjectId}; creating a separate document set.",
                    project.HubSpotLink, project.ProjectId, scoping.ProjectId);
            }
            else if (scoping is not null)
            {
                return await PromoteScopingWorkspaceAsync(ctx, scoping.Url, project, destination, cancellationToken);
            }
            else
            {
                // Surfaces a mistyped or stale value: this engagement gets a new folder below, and any
                // scoping folder it should have inherited stays orphaned for a human to merge.
                _logger.LogWarning(
                    "Project {ProjectId} carries HubSpot link {Link} but no scoping workspace matched it; creating a new document set.",
                    project.ProjectId, project.HubSpotLink);
            }
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
        var newSig = ReconcileSignature.Compute(project, destination.PracticeLeaderEmail, destination.AdminEmails, destination.MembersPermissionLevel);
        await ApplyMetadataAsync(ctx, serverRelativeUrl, project, newSig, siteUrl, cancellationToken);
        await ApplyPermissionsAsync(ctx, serverRelativeUrl, project, destination, cancellationToken);

        // Client uploads folder + external "Request files" link. Create-once: only on first creation,
        // after permissions (so the child-scoped sharing link isn't cleared by the inheritance break).
        string? uploadLink = null;
        if (_options.CreateClientUploadLink)
        {
            uploadLink = await EnsureClientUploadsAsync(ctx, list, serverRelativeUrl, siteUrl, cancellationToken);
        }

        await _notifier.NotifyCreatedAsync(
            ProjectNotice(project, destination, siteUrl, serverRelativeUrl, uploadLink),
            ProjectRecipients(project, destination),
            WorkspaceNotifier.LeaderToExclude(destination.PracticeLeaderEmail, ProjectLeadPeople(project)),
            cancellationToken);

        return new DocumentSetResult(Created: true, serverRelativeUrl,
            DataroomUrl: BuildAbsoluteUrl(siteUrl, serverRelativeUrl), ClientUploadUrl: uploadLink);
    }

    public async Task<DocumentSetResult> EnsureScopingWorkspaceAsync(
        ScopingWorkspace workspace, CancellationToken cancellationToken)
    {
        var destination = ResolveDestination(workspace.Practice);
        var siteUrl = string.IsNullOrWhiteSpace(destination.SiteUrl) ? _options.SiteUrl : destination.SiteUrl!;

        using var ctx = await _contextFactory.CreateContextAsync(siteUrl);
        var list = ctx.Web.Lists.GetByTitle(destination.Library);
        ctx.Load(list, l => l.RootFolder.ServerRelativeUrl, l => l.ContentTypes);
        await ctx.ExecuteQueryRetryAsync();

        await EnsureTextColumnAsync(ctx, list, _options.HubSpotDealIdColumn);
        await EnsureTextColumnAsync(ctx, list, _options.OpportunityIdColumn);
        await EnsureTextColumnAsync(ctx, list, _options.StatusColumn);

        // Idempotency keyed on the HubSpot deal id — immutable, unlike the opportunity number, which can
        // be assigned or corrected later. Keying on the id is what stops a late-arriving opportunity number
        // from looking like a new engagement and producing a duplicate workspace.
        var existing = await FindByColumnAsync(ctx, list, destination.ParentFolder, _options.HubSpotDealIdColumn, workspace.DealId, cancellationToken);
        if (existing is not null && !string.IsNullOrWhiteSpace(existing.ProjectId))
        {
            // The workspace has been promoted: it is an Acumatica project now, and Acumatica owns its
            // metadata and permissions. Re-applying the scoping state here would flip Status back to
            // Scoping and reset access to the deal owner, wiping the delivery team — so leave it alone.
            _logger.LogInformation(
                "Deal {DealId} was promoted to project {ProjectId}; leaving the workspace to the Acumatica sync.",
                workspace.DealId, existing.ProjectId);
            return new DocumentSetResult(Created: false, existing.Url);
        }

        if (existing is not null)
        {
            _logger.LogInformation("Scoping workspace for deal {DealId} already exists at {Url}; updating.", workspace.DealId, existing.Url);

            // Self-heal the folder name: the customer can change in HubSpot after the room was created, so
            // the sync is authoritative — re-derive the name each poll and rename in place. (Rename first;
            // it changes the server-relative URL that metadata + permissions then address.)
            var (desired, discriminator) = ScopingName(workspace);
            var url = await TryRenameDocumentSetAsync(ctx, existing.Url, desired, discriminator, cancellationToken);

            await ApplyScopingMetadataAsync(ctx, url, workspace, cancellationToken);
            var addedScoping = await ApplyScopingPermissionsAsync(ctx, url, workspace, destination, cancellationToken);
            if (addedScoping.Count > 0)
            {
                await _notifier.NotifyAccessAddedAsync(
                    ScopingNotice(workspace, destination, siteUrl, url, await ReadUploadLinkAsync(ctx, url)),
                    addedScoping, WorkspaceNotifier.LeaderToExclude(destination.PracticeLeaderEmail, ScopingLeadPeople(workspace)),
                    cancellationToken);
            }
            return new DocumentSetResult(Created: false, url);
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

        // Name on the opportunity number (PQCode) so the folder reads the way people refer to the
        // engagement — and matches what gets typed into Acumatica at conversion.
        var (desiredName, nameId) = ScopingName(workspace);
        var setName = await ResolveUniqueNameAsync(ctx, parentFolder, desiredName, nameId, cancellationToken);

        _logger.LogInformation("Creating scoping workspace '{Name}' for opportunity {OpportunityId} (deal {DealId}, practice '{Practice}').",
            setName, workspace.OpportunityId ?? "<none>", workspace.DealId, workspace.Practice ?? "<none>");

        var created = SPDocumentSet.Create(ctx, parentFolder, setName, contentType.Id);
        await ctx.ExecuteQueryRetryAsync();
        var serverRelativeUrl = created.Value;

        await ApplyScopingMetadataAsync(ctx, serverRelativeUrl, workspace, cancellationToken);
        await ApplyScopingPermissionsAsync(ctx, serverRelativeUrl, workspace, destination, cancellationToken);
        string? scopingUploadLink = null;
        if (_options.CreateClientUploadLink)
        {
            scopingUploadLink = await EnsureClientUploadsAsync(ctx, list, serverRelativeUrl, siteUrl, cancellationToken);
        }

        await _notifier.NotifyCreatedAsync(
            ScopingNotice(workspace, destination, siteUrl, serverRelativeUrl, scopingUploadLink),
            ScopingRecipients(workspace, destination),
            WorkspaceNotifier.LeaderToExclude(destination.PracticeLeaderEmail, ScopingLeadPeople(workspace)),
            cancellationToken);

        return new DocumentSetResult(Created: true, serverRelativeUrl);
    }

    /// <summary>
    /// The desired folder name for a scoping workspace, plus the discriminator used to keep it unique.
    /// Named on the opportunity number (PQCode) so it reads the way people refer to the engagement and
    /// matches what is typed into Acumatica at conversion; falls back to the deal record id for older deals
    /// that predate the sequential opportunity number.
    /// </summary>
    private (string DesiredName, string Discriminator) ScopingName(ScopingWorkspace workspace)
    {
        var discriminator = string.IsNullOrWhiteSpace(workspace.OpportunityId) ? workspace.DealId : workspace.OpportunityId!;
        var basis = workspace.CustomerName ?? workspace.ProjectName ?? discriminator;
        return (SharePointNaming.BuildDocumentSetName(basis, discriminator, _options.DocumentSetNameMaxLength), discriminator);
    }

    // ----- Notification helpers -----

    // The roles that mean the practice leader is running this engagement themselves: PM or principal in
    // charge for a project, deal owner or PM for a scoping deal. A leader in one of those roles gets the
    // emails; a leader who's only on the team, an originator, or on the room because they lead the
    // practice is left out.
    private static IEnumerable<string?> ProjectLeadPeople(AcumaticaProject project)
    {
        yield return project.ProjectManagerEmail;
        yield return project.PrincipalInChargeEmail;
    }

    private static IEnumerable<string?> ScopingLeadPeople(ScopingWorkspace ws)
    {
        yield return ws.OwnerEmail;
        yield return ws.ProjectManagerEmail;
    }

    private static IEnumerable<string?> ProjectRecipients(AcumaticaProject project, PracticeMappingEntry destination)
    {
        yield return project.ProjectManagerEmail;
        yield return project.PrincipalInChargeEmail;
        foreach (var e in project.TeamEmails) yield return e;
        foreach (var e in destination.AdminEmails) yield return e;
    }

    private static IEnumerable<string?> ScopingRecipients(ScopingWorkspace ws, PracticeMappingEntry destination)
    {
        yield return ws.OwnerEmail;
        yield return ws.ProjectManagerEmail;
        yield return ws.OriginatorAEmail;
        yield return ws.OriginatorBEmail;
        foreach (var e in destination.AdminEmails) yield return e;
    }

    // Email of a People column value (or a plain-text column that holds an address).
    private static string? PersonEmail(ListItem item, string column)
        => item.FieldValues.TryGetValue(column, out var v) ? v switch
        {
            FieldUserValue u => u.Email,
            string text when text.Contains('@') => text,
            _ => null,
        } : null;

    /// <summary>The practice label to show people: the destination's DisplayName override, else the raw value.</summary>
    private static string? PracticeLabel(PracticeMappingEntry destination, string? rawPractice)
        => !string.IsNullOrWhiteSpace(destination.DisplayName) ? destination.DisplayName : rawPractice;

    private WorkspaceNotice ProjectNotice(AcumaticaProject p, PracticeMappingEntry destination, string siteUrl, string serverRelativeUrl, string? uploadLink) => new()
    {
        Phase = WorkspacePhase.Execution,
        CustomerName = string.IsNullOrWhiteSpace(p.CustomerName) ? p.ProjectId : p.CustomerName!,
        EngagementName = p.ProjectName,
        IdLabel = "Project ID",
        IdValue = p.ProjectId,
        ProjectManager = p.ProjectManager,
        Practice = PracticeLabel(destination, p.Practice),
        DataroomUrl = BuildAbsoluteUrl(siteUrl, serverRelativeUrl),
        UploadLinkUrl = uploadLink,
    };

    private WorkspaceNotice ScopingNotice(ScopingWorkspace w, PracticeMappingEntry destination, string siteUrl, string serverRelativeUrl, string? uploadLink) => new()
    {
        Phase = WorkspacePhase.Scoping,
        CustomerName = w.CustomerName ?? w.OpportunityId ?? w.DealId,
        EngagementName = w.ProjectName,
        IdLabel = "Opportunity #",
        IdValue = w.OpportunityId ?? w.DealId,
        Practice = PracticeLabel(destination, w.Practice),
        DataroomUrl = BuildAbsoluteUrl(siteUrl, serverRelativeUrl),
        UploadLinkUrl = uploadLink,
    };

    /// <summary>Reads the stored client upload link off a document set (for access-added emails).</summary>
    private async Task<string?> ReadUploadLinkAsync(ClientContext ctx, string serverRelativeUrl)
    {
        var col = _options.ClientUploadLinkColumn;
        if (string.IsNullOrWhiteSpace(col))
        {
            return null;
        }

        try
        {
            var item = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl)).ListItemAllFields;
            ctx.Load(item);
            await ctx.ExecuteQueryRetryAsync();
            return item.FieldValues.TryGetValue(col, out var v) ? v?.ToString() : null;
        }
        catch
        {
            return null;
        }
    }

    private async Task NotifyProjectAccessAddedAsync(
        ClientContext ctx, AcumaticaProject project, PracticeMappingEntry destination,
        string serverRelativeUrl, string siteUrl, IReadOnlyList<string> newlyAdded, CancellationToken cancellationToken)
    {
        if (newlyAdded.Count == 0)
        {
            return;
        }

        var uploadLink = await ReadUploadLinkAsync(ctx, serverRelativeUrl);
        await _notifier.NotifyAccessAddedAsync(
            ProjectNotice(project, destination, siteUrl, serverRelativeUrl, uploadLink),
            newlyAdded, WorkspaceNotifier.LeaderToExclude(destination.PracticeLeaderEmail, ProjectLeadPeople(project)),
            cancellationToken);
    }

    // Delegates to the tested helper. DocumentSet.Create (and some other CSOM calls) can hand back an
    // ABSOLUTE url; the helper normalizes it to a server-relative path first so the host isn't doubled
    // (…sharepoint.com…sharepoint.com/sites/…) — the bug that broke "Open the dataroom" in creation emails.
    private static string BuildAbsoluteUrl(string siteUrl, string serverRelativeUrl)
        => SharePointNaming.ToAbsoluteUrl(siteUrl, serverRelativeUrl);

    /// <summary>
    /// Finds files added to any "Client Uploads" folder since <paramref name="since"/> and emails everyone
    /// with access to that workspace (the practice leader is kept for scoping rooms, dropped for engagements).
    /// The email states how many files and lists their names. Fail-soft per workspace.
    /// </summary>
    public async Task<ClientUploadScanResult> ScanAndNotifyClientUploadsAsync(
        DateTimeOffset since, CancellationToken cancellationToken)
    {
        var result = new ClientUploadScanResult();

        foreach (var mapping in _options.PracticeMappings)
        {
            // One site failing must not block the others — or the caller's watermark. On 2026-09-29 the ESOP
            // library passed the 5,000-item list view threshold, the scan threw on that site, the watermark
            // never advanced, and every earlier site re-sent the same upload emails every 15 minutes.
            try
            {
                await ScanSiteClientUploadsAsync(mapping, since, result, cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                result.SitesFailed++;
                _logger.LogError(ex, "Client-upload scan failed for practice {Practice}; other sites continue. Uploads there since {Since:o} were not notified.",
                    mapping.Practice, since);
            }
        }

        _logger.LogInformation(
            "Client-upload scan since {Since:o}: {Files} new file(s) across {Workspaces} workspace(s); {Notified} notified; {Failed} site(s) failed.",
            since, result.NewFiles, result.WorkspacesWithNewFiles, result.Notified, result.SitesFailed);
        return result;
    }

    private async Task ScanSiteClientUploadsAsync(
        PracticeMappingEntry mapping, DateTimeOffset since, ClientUploadScanResult result, CancellationToken cancellationToken)
    {
        var marker = "/" + _options.ClientUploadsFolderName.Trim('/') + "/";
        {
            var siteUrl = string.IsNullOrWhiteSpace(mapping.SiteUrl) ? _options.SiteUrl : mapping.SiteUrl!;
            using var ctx = await _contextFactory.CreateContextAsync(siteUrl);
            var list = ctx.Web.Lists.GetByTitle(mapping.Library);
            ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
            await ctx.ExecuteQueryRetryAsync();
            await EnsureCreatedIndexedAsync(ctx, list, siteUrl);

            // Files created since the watermark under a Client Uploads folder, grouped by document set.
            // The library-wide query is cheapest; when SharePoint refuses it (library over the list view
            // threshold without a Created index — e.g. ESOP after the MN migration — or a copy-migration
            // that creates 5,000+ files in one window), walk each room's Client Uploads folder instead.
            Dictionary<string, List<string>> newByDocSet;
            try
            {
                newByDocSet = await FindNewUploadsByQueryAsync(ctx, list, since, marker);
            }
            catch (ServerException ex) when (ClientUploadQuery.IsListViewThreshold(ex.ServerErrorTypeName, ex.Message))
            {
                _logger.LogWarning("Client-upload query on {Site} exceeds the list view threshold; checking each room's {Folder} folder instead.",
                    siteUrl, _options.ClientUploadsFolderName);
                newByDocSet = await FindNewUploadsByFolderWalkAsync(ctx, list, mapping.ParentFolder, since);
            }

            foreach (var (docSetUrl, fileNames) in newByDocSet)
            {
                cancellationToken.ThrowIfCancellationRequested();
                result.NewFiles += fileNames.Count;
                result.WorkspacesWithNewFiles++;

                try
                {
                    var item = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(docSetUrl)).ListItemAllFields;
                    ctx.Load(item,
                        i => i[_options.CustomerNameColumn], i => i[_options.ProjectNameColumn],
                        i => i[_options.ProjectIdColumn], i => i[_options.OpportunityIdColumn],
                        i => i[_options.StatusColumn], i => i[_options.ClientUploadLinkColumn],
                        i => i[_options.ProjectManagerColumn]);
                    ctx.Load(item.RoleAssignments, r => r.Include(a => a.Member.PrincipalType, a => a.Member.LoginName));
                    await ctx.ExecuteQueryRetryAsync();

                    string? Field(string col) => item.FieldValues.TryGetValue(col, out var v) ? v?.ToString() : null;
                    var status = Field(_options.StatusColumn);
                    var isScoping = string.Equals(status, _options.ScopingStatusValue, StringComparison.OrdinalIgnoreCase);

                    var recipients = GranteeEmails(item.RoleAssignments);
                    // Keep the practice leader for scoping rooms; drop for engagements (not delivering work)
                    // unless they're the engagement's PM. The PIC isn't stored on the item, so a leader who
                    // is only the PIC is still dropped from upload emails.
                    var exclude = isScoping ? null
                        : WorkspaceNotifier.LeaderToExclude(mapping.PracticeLeaderEmail, new[] { PersonEmail(item, _options.ProjectManagerColumn) });

                    var notice = new WorkspaceNotice
                    {
                        Phase = isScoping ? WorkspacePhase.Scoping : WorkspacePhase.Execution,
                        CustomerName = Field(_options.CustomerNameColumn) ?? "(unknown)",
                        EngagementName = Field(_options.ProjectNameColumn),
                        IdLabel = isScoping ? "Opportunity #" : "Project ID",
                        IdValue = isScoping ? Field(_options.OpportunityIdColumn) : Field(_options.ProjectIdColumn),
                        Practice = PracticeLabel(mapping, mapping.Practice),
                        DataroomUrl = BuildAbsoluteUrl(siteUrl, docSetUrl),
                        UploadLinkUrl = Field(_options.ClientUploadLinkColumn),
                    };
                    var uploadsFolderUrl = BuildAbsoluteUrl(siteUrl, docSetUrl + "/" + _options.ClientUploadsFolderName);

                    await _notifier.NotifyClientUploadAsync(
                        notice, fileNames, uploadsFolderUrl, recipients, exclude, cancellationToken);
                    result.Notified++;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to notify client uploads for {Url}; continuing.", docSetUrl);
                }
            }
        }
    }

    /// <summary>Library-wide CAML query for new files, grouped by owning document set.</summary>
    private static async Task<Dictionary<string, List<string>>> FindNewUploadsByQueryAsync(
        ClientContext ctx, List list, DateTimeOffset since, string marker)
    {
        var found = new List<(string FileRef, string Leaf)>();
        ListItemCollectionPosition? position = null;
        do
        {
            var query = new CamlQuery
            {
                ViewXml = ClientUploadQuery.Build(since),
                ListItemCollectionPosition = position,
            };
            var items = list.GetItems(query);
            ctx.Load(items, c => c.ListItemCollectionPosition, c => c.Include(i => i["FileRef"], i => i["FileLeafRef"]));
            await ctx.ExecuteQueryRetryAsync();

            foreach (var it in items)
            {
                var fileRef = it["FileRef"]?.ToString();
                var leaf = it["FileLeafRef"]?.ToString();
                if (!string.IsNullOrWhiteSpace(fileRef) && !string.IsNullOrWhiteSpace(leaf))
                {
                    found.Add((fileRef, leaf));
                }
            }

            position = items.ListItemCollectionPosition;
        }
        while (position is not null);

        return ClientUploadQuery.GroupByDocSet(found, marker);
    }

    /// <summary>
    /// Threshold-proof fallback: lists each room under the practice's parent folder, then walks that room's
    /// Client Uploads folder with the Files/Folders collections (not subject to the list view threshold) and
    /// keeps files created after <paramref name="since"/>. Folders under the parent with no Client Uploads
    /// folder (e.g. ESOP's migrated "MN ESOP Clients") are skipped.
    /// </summary>
    private async Task<Dictionary<string, List<string>>> FindNewUploadsByFolderWalkAsync(
        ClientContext ctx, List list, string? parentFolder, DateTimeOffset since)
    {
        var root = list.RootFolder.ServerRelativeUrl.TrimEnd('/');
        var parentUrl = string.IsNullOrWhiteSpace(parentFolder) ? root : $"{root}/{parentFolder.Trim('/')}";
        var parent = ctx.Web.GetFolderByServerRelativeUrl(parentUrl);
        // Rooms and their immediate subfolders in one round trip. (Loading Exists on a missing folder
        // throws "File Not Found" in CSOM, so find Client Uploads by name instead.)
        ctx.Load(parent.Folders, fs => fs.Include(
            f => f.Name,
            f => f.Folders.Include(c => c.Name, c => c.ServerRelativeUrl)));
        await ctx.ExecuteQueryRetryAsync();

        var sinceUtc = since.UtcDateTime;
        var found = new List<(string FileRef, string Leaf)>();
        foreach (var room in parent.Folders)
        {
            var uploads = room.Folders.FirstOrDefault(c =>
                c.Name.Equals(_options.ClientUploadsFolderName, StringComparison.OrdinalIgnoreCase));
            if (uploads is null)
            {
                continue; // not a room (e.g. "MN ESOP Clients") or a room without an uploads folder
            }

            await CollectFilesCreatedAfterAsync(ctx, ctx.Web.GetFolderByServerRelativeUrl(uploads.ServerRelativeUrl), sinceUtc, found);
        }

        return ClientUploadQuery.GroupByDocSet(found, "/" + _options.ClientUploadsFolderName.Trim('/') + "/");
    }

    private static async Task CollectFilesCreatedAfterAsync(
        ClientContext ctx, Folder folder, DateTime sinceUtc, List<(string FileRef, string Leaf)> found)
    {
        ctx.Load(folder.Files, fs => fs.Include(f => f.Name, f => f.ServerRelativeUrl, f => f.TimeCreated));
        ctx.Load(folder.Folders, fs => fs.Include(f => f.Name, f => f.ServerRelativeUrl));
        await ctx.ExecuteQueryRetryAsync();

        foreach (var file in folder.Files)
        {
            if (file.TimeCreated > sinceUtc)
            {
                found.Add((file.ServerRelativeUrl, file.Name));
            }
        }

        foreach (var sub in folder.Folders)
        {
            await CollectFilesCreatedAfterAsync(ctx, sub, sinceUtc, found);
        }
    }

    // Sites whose Created column is known to be indexed (checked once per host process).
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, bool> CreatedIndexedSites =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Indexes the library's Created column so the scan's Created filter stays under the list view
    /// threshold once a library passes 5,000 items. Fail-soft: without it the query still works on small
    /// libraries, and a large one fails that site alone.
    /// </summary>
    private async Task EnsureCreatedIndexedAsync(ClientContext ctx, List list, string siteUrl)
    {
        if (CreatedIndexedSites.ContainsKey(siteUrl))
        {
            return;
        }

        try
        {
            var created = list.Fields.GetByInternalNameOrTitle("Created");
            ctx.Load(created, f => f.Indexed);
            await ctx.ExecuteQueryRetryAsync();
            if (!created.Indexed)
            {
                created.Indexed = true;
                created.Update();
                await ctx.ExecuteQueryRetryAsync();
                _logger.LogInformation("Indexed the Created column on {Site} for the client-upload scan.", siteUrl);
            }
            CreatedIndexedSites[siteUrl] = true;
        }
        catch (Exception ex)
        {
            // SharePoint refuses to index a library already past the threshold. Don't retry every cycle —
            // the folder-walk fallback covers it; the next host restart tries again.
            CreatedIndexedSites[siteUrl] = false;
            _logger.LogWarning("Couldn't index the Created column on {Site}: {Message}", siteUrl, ex.Message);
        }
    }

    /// <summary>Resolves a document set's role assignments to the email addresses of its individual members.</summary>
    private static IReadOnlyList<string> GranteeEmails(RoleAssignmentCollection roleAssignments)
    {
        var emails = new List<string>();
        foreach (var ra in roleAssignments)
        {
            // Only individual users — skip the Owners group and the app principal.
            if (ra.Member.PrincipalType != Microsoft.SharePoint.Client.Utilities.PrincipalType.User)
            {
                continue;
            }

            // Claims login looks like "i:0#.f|membership|user@domain"; the UPN/email is the trailing segment.
            var login = ra.Member.LoginName ?? string.Empty;
            var candidate = login.Split('|').LastOrDefault();
            if (!string.IsNullOrWhiteSpace(candidate) && candidate.Contains('@'))
            {
                emails.Add(candidate.Trim());
            }
        }

        return emails.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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

    public async Task<ReconcileResult> ReconcileAsync(
        IReadOnlyList<AcumaticaProject> desiredProjects,
        IReadOnlySet<string>? onlyProjectIds,
        CancellationToken cancellationToken)
    {
        int considered = 0, updated = 0, unchanged = 0, notTracked = 0;
        var resyncs = new List<UrlResync>();

        // One SharePoint context per (site, library) — currently a single group (Estate & Gift).
        var groups = desiredProjects
            .Select(p => (Project: p, Dest: ResolveDestination(p.Practice)))
            .GroupBy(x => (
                Site: string.IsNullOrWhiteSpace(x.Dest.SiteUrl) ? _options.SiteUrl : x.Dest.SiteUrl!,
                x.Dest.Library));

        var failedSites = new List<string>();
        foreach (var group in groups)
        {
            // One site failing (throttling, an outage, a permissions change) must not stop the others.
            try
            {
                using var ctx = await _contextFactory.CreateContextAsync(group.Key.Site);
                var list = ctx.Web.Lists.GetByTitle(group.Key.Library);
                ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
                await ctx.ExecuteQueryRetryAsync();

                await EnsureSignatureColumnAsync(ctx, list);
                await EnsureTextColumnAsync(ctx, list, _options.StatusColumn);
                await EnsureTextColumnAsync(ctx, list, _options.DataroomUrlColumn);
                var parentFolders = group.Select(x => x.Dest.ParentFolder).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var tracked = await GetTrackedDocSetsAsync(ctx, list, parentFolders, cancellationToken);

                foreach (var x in group)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (onlyProjectIds is not null && !onlyProjectIds.Contains(x.Project.ProjectId))
                    {
                        continue;
                    }

                    if (!tracked.TryGetValue(x.Project.ProjectId, out var t))
                    {
                        notTracked++; // untracked project — no backfill
                        continue;
                    }

                    considered++;

                    // Detect a manual folder rename: the folder's current URL no longer matches the one stamped
                    // in the metadata column. (A rename touches nothing in Acumatica, so signature-gating alone
                    // would never catch it.) An empty column means the workspace predates this feature — seed it
                    // silently so we don't fire a write-back storm on the first sweep.
                    var currentUrl = BuildAbsoluteUrl(group.Key.Site, t.Url);
                    var urlMissing = string.IsNullOrWhiteSpace(t.DataroomUrl);
                    var urlDrifted = !urlMissing && !string.Equals(t.DataroomUrl, currentUrl, StringComparison.OrdinalIgnoreCase);

                    var desiredSig = ReconcileSignature.Compute(x.Project, x.Dest.PracticeLeaderEmail, x.Dest.AdminEmails, x.Dest.MembersPermissionLevel);
                    var sigChanged = !string.Equals(desiredSig, t.Signature, StringComparison.OrdinalIgnoreCase);

                    if (sigChanged)
                    {
                        _logger.LogInformation("Reconcile: project {ProjectId} changed — re-applying metadata + permissions.", x.Project.ProjectId);
                        // ApplyMetadataAsync also (re)stamps the dataroom URL column to the current URL.
                        await ApplyMetadataAsync(ctx, t.Url, x.Project, desiredSig, group.Key.Site, cancellationToken);
                        var reconcileAdded = await ApplyPermissionsAsync(ctx, t.Url, x.Project, x.Dest, cancellationToken);
                        await NotifyProjectAccessAddedAsync(ctx, x.Project, x.Dest, t.Url, group.Key.Site, reconcileAdded, cancellationToken);
                        updated++;
                    }
                    else if (urlDrifted)
                    {
                        _logger.LogInformation("Reconcile: project {ProjectId} folder was renamed — refreshing dataroom URL in SharePoint + Acumatica.", x.Project.ProjectId);
                        await StampDataroomUrlAsync(ctx, t.Url, currentUrl, cancellationToken);
                        updated++;
                    }
                    else if (urlMissing)
                    {
                        await StampDataroomUrlAsync(ctx, t.Url, currentUrl, cancellationToken); // seed baseline, no write-back
                        unchanged++;
                    }
                    else
                    {
                        unchanged++;
                    }

                    // Re-write Acumatica's DATAURL only for a genuine rename (a prior value existed and changed).
                    if (urlDrifted)
                    {
                        resyncs.Add(new UrlResync(x.Project.ProjectId, currentUrl));
                    }
                }
            }
            catch (Exception ex) when (!cancellationToken.IsCancellationRequested)
            {
                failedSites.Add(group.Key.Site);
                _logger.LogError(ex, "Reconcile failed for {Site} ({Library}); continuing with the other sites.",
                    group.Key.Site, group.Key.Library);
            }
        }

        if (failedSites.Count > 0)
        {
            _logger.LogError("Reconcile finished with {Count} failed site(s): {Sites}.", failedSites.Count, string.Join(", ", failedSites));
        }

        return new ReconcileResult
        {
            Considered = considered,
            Updated = updated,
            Unchanged = unchanged,
            NotTracked = notTracked,
            UrlResyncs = resyncs,
        };
    }

    /// <summary>Refreshes only the dataroom-URL metadata column on a document set (used when a folder was renamed).</summary>
    private async Task StampDataroomUrlAsync(ClientContext ctx, string serverRelativeUrl, string dataroomUrl, CancellationToken cancellationToken)
    {
        var item = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl)).ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();
        SetIfPresent(item, _options.DataroomUrlColumn, dataroomUrl);
        item.Update();
        await ctx.ExecuteQueryRetryAsync();
    }

    /// <summary>Ensures the hidden signature column exists on the library (idempotent).</summary>
    private async Task EnsureSignatureColumnAsync(ClientContext ctx, List list)
    {
        var col = _options.SignatureColumn;
        if (string.IsNullOrWhiteSpace(col))
        {
            return;
        }

        ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName));
        await ctx.ExecuteQueryRetryAsync();
        if (list.Fields.Any(f => f.InternalName == col))
        {
            return;
        }

        list.Fields.AddFieldAsXml(
            $"<Field Type='Text' Name='{col}' StaticName='{col}' DisplayName='{col}' Hidden='TRUE' Group='ProjectSync'/>",
            addToDefaultView: false, options: AddFieldOptions.AddFieldInternalNameHint);
        await ctx.ExecuteQueryRetryAsync();
        _logger.LogInformation("Created reconcile signature column '{Column}'.", col);
    }

    /// <summary>
    /// Reads every tracked room (one with a Project Id) under the given parent folders: project id ->
    /// (folder url, stored signature, stored dataroom url). Reads only the parent folders' direct children
    /// (see <see cref="RoomQuery"/>), so it works however many files the library holds.
    /// </summary>
    private async Task<Dictionary<string, (string Url, string? Signature, string? DataroomUrl)>> GetTrackedDocSetsAsync(
        ClientContext ctx, List list, IReadOnlyList<string?> parentFolders, CancellationToken cancellationToken)
    {
        var result = new Dictionary<string, (string, string?, string?)>(StringComparer.OrdinalIgnoreCase);
        var pidCol = _options.ProjectIdColumn;
        var sigCol = _options.SignatureColumn;
        var urlCol = _options.DataroomUrlColumn;
        foreach (var parent in parentFolders)
        {
            var rooms = await ReadRoomsAsync(ctx, list, parent, new[] { pidCol, sigCol, urlCol }, cancellationToken);
            foreach (var it in rooms)
            {
                var pid = Field(it, pidCol)?.Trim();
                if (string.IsNullOrWhiteSpace(pid))
                {
                    continue;
                }

                result[pid!] = (Field(it, "FileRef") ?? string.Empty, Field(it, sigCol), Field(it, urlCol));
            }
        }

        return result;
    }

    /// <summary>
    /// All rooms that are direct children of <paramref name="parentFolder"/> (the mapping's ParentFolder,
    /// e.g. Projects/Current), with the requested fields plus FileRef. Unfiltered, ID-ordered paging scoped
    /// to that folder, so the list view threshold never applies.
    /// </summary>
    private async Task<List<ListItem>> ReadRoomsAsync(
        ClientContext ctx, List list, string? parentFolder, IReadOnlyList<string> fields, CancellationToken cancellationToken)
    {
        ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
        await ctx.ExecuteQueryRetryAsync();
        var folderUrl = RoomQuery.ParentFolderUrl(list.RootFolder.ServerRelativeUrl, parentFolder);

        var wanted = fields.Where(f => !string.IsNullOrWhiteSpace(f)).Append("FileRef").Distinct().ToList();
        var rooms = new List<ListItem>();
        ListItemCollectionPosition? position = null;
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            var query = new CamlQuery
            {
                ViewXml = RoomQuery.FolderChildren(wanted),
                FolderServerRelativeUrl = folderUrl,
                ListItemCollectionPosition = position,
            };
            var items = list.GetItems(query);
            ctx.Load(items, c => c.ListItemCollectionPosition);
            ctx.Load(items, c => c.Include(i => i["FSObjType"], i => i["ContentTypeId"]));
            foreach (var f in wanted)
            {
                var name = f;
                ctx.Load(items, c => c.Include(i => i[name]));
            }
            await ctx.ExecuteQueryRetryAsync();

            rooms.AddRange(items.Where(i => RoomQuery.IsRoom(i["FSObjType"], i["ContentTypeId"])));
            position = items.ListItemCollectionPosition;
        }
        while (position is not null);

        return rooms;
    }

    /// <summary>
    /// Rooms whose <paramref name="column"/> equals <paramref name="value"/> (at most <paramref name="max"/>).
    /// Tries the indexed lookup first (instant at any size, and finds the room anywhere in the library); if
    /// SharePoint still refuses it (the column isn't indexed yet, or the library is too large to index),
    /// falls back to reading the parent folder's rooms and matching in code.
    /// </summary>
    private async Task<List<ListItem>> FindRoomsByColumnAsync(
        ClientContext ctx, List list, string? parentFolder, string column, string value, int max,
        IReadOnlyList<string> fields, CancellationToken cancellationToken)
    {
        await EnsureLookupIndexesAsync(ctx, list);
        var wanted = fields.Append("FileRef").Append(column).Distinct().ToList();
        try
        {
            var items = list.GetItems(new CamlQuery { ViewXml = RoomQuery.Lookup(column, value, max, wanted) });
            ctx.Load(items, c => c.Include(i => i.FileSystemObjectType));
            foreach (var f in wanted)
            {
                var name = f;
                ctx.Load(items, c => c.Include(i => i[name]));
            }
            await ctx.ExecuteQueryRetryAsync();
            return items.ToList();
        }
        catch (ServerException ex) when (ClientUploadQuery.IsListViewThreshold(ex.ServerErrorTypeName, ex.Message))
        {
            _logger.LogInformation("Lookup on {Column} exceeds the list view threshold; matching the rooms in {Folder} instead.",
                column, parentFolder ?? "<library root>");
            var rooms = await ReadRoomsAsync(ctx, list, parentFolder, wanted, cancellationToken);
            return rooms
                .Where(r => string.Equals(Field(r, column)?.Trim(), value.Trim(), StringComparison.OrdinalIgnoreCase))
                .Take(max)
                .ToList();
        }
    }

    // Only index while comfortably under the 5,000-item list view threshold, where it's instant.
    private const int IndexableItemLimit = 4500;

    // Libraries whose lookup columns were already checked for an index in this process.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<Guid, bool> _lookupIndexesChecked = new();

    /// <summary>
    /// Indexes the room key columns (Project Id, HubSpot deal id, opportunity number) while the library is
    /// still under the list view threshold, so lookups on them stay allowed as it grows. Once per library per
    /// process. A library already over the threshold is left alone: SharePoint builds the index inside the
    /// request there, which can hang for minutes (it stalled the scoping poll past its 5-minute limit on
    /// Commercial Litigation), and the lookup's folder-read fallback handles those libraries anyway.
    /// </summary>
    private async Task EnsureLookupIndexesAsync(ClientContext ctx, List list)
    {
        ctx.Load(list, l => l.Id, l => l.ItemCount);
        await ctx.ExecuteQueryRetryAsync();
        if (_lookupIndexesChecked.ContainsKey(list.Id))
        {
            return;
        }

        if (list.ItemCount >= IndexableItemLimit)
        {
            _lookupIndexesChecked.TryAdd(list.Id, true);
            return;
        }

        var keys = new[] { _options.ProjectIdColumn, _options.HubSpotDealIdColumn, _options.OpportunityIdColumn }
            .Where(c => !string.IsNullOrWhiteSpace(c)).ToList();
        ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName, f => f.Indexed));
        await ctx.ExecuteQueryRetryAsync();
        foreach (var field in list.Fields.Where(f => keys.Contains(f.InternalName) && !f.Indexed).ToList())
        {
            try
            {
                field.Indexed = true;
                field.Update();
                await ctx.ExecuteQueryRetryAsync();
                _logger.LogInformation("Indexed column '{Column}' for threshold-safe lookups.", field.InternalName);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Couldn't index column '{Column}' ({Message}); lookups will fall back to reading the rooms folder.",
                    field.InternalName, ex.Message);
            }
        }

        _lookupIndexesChecked.TryAdd(list.Id, true);
    }

    private static string? Field(ListItem item, string column) =>
        item.FieldValues.TryGetValue(column, out var v) ? v?.ToString() : null;

    private PracticeMappingEntry ResolveDestination(string? practice)
    {
        var key = practice?.Trim();
        if (!string.IsNullOrEmpty(key))
        {
            // A HubSpot deal's practice is a multi-select, so it arrives as a ';'-delimited list
            // (e.g. "Estate & Gift;Tangible Assets"). The include-filter upstream is a contains-match, so
            // match a mapping if ANY token matches — otherwise a legitimately in-scope multi-practice deal
            // fails to resolve a destination and (with no '*' fallback) throws, jamming the poll.
            var tokens = key
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            var match = _options.PracticeMappings.FirstOrDefault(m =>
                string.Equals(m.Practice.Trim(), key, StringComparison.OrdinalIgnoreCase) ||
                tokens.Any(t => string.Equals(m.Practice.Trim(), t, StringComparison.OrdinalIgnoreCase)));
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
        ClientContext ctx, List list, string? parentFolder, string projectId, CancellationToken cancellationToken)
    {
        var rooms = await FindRoomsByColumnAsync(ctx, list, parentFolder, _options.ProjectIdColumn, projectId, 1,
            Array.Empty<string>(), cancellationToken);
        return rooms.Count > 0 ? Field(rooms[0], "FileRef") : null;
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

    /// <summary>
    /// Promotes an existing scoping workspace (born from a HubSpot deal) into the execution phase IN PLACE:
    /// renames it to the project-id form, stamps the Acumatica metadata — which sets Project Id and flips
    /// Status to Execution — and re-applies permissions for the delivery team. Nothing moves: the folder,
    /// its documents, and its client-upload link stay exactly where the scoping phase left them.
    /// </summary>
    private async Task<DocumentSetResult> PromoteScopingWorkspaceAsync(
        ClientContext ctx, string scopingUrl, AcumaticaProject project, PracticeMappingEntry destination, CancellationToken cancellationToken)
    {
        _logger.LogInformation(
            "Promoting scoping workspace {Url} to project {ProjectId} (HubSpot link {Link}).",
            scopingUrl, project.ProjectId, project.HubSpotLink);

        // Rename FIRST: it changes the server-relative URL, so everything after must use the new one.
        var desiredName = SharePointNaming.BuildDocumentSetName(
            project.CustomerName, project.ProjectId, _options.DocumentSetNameMaxLength);
        var url = await TryRenameDocumentSetAsync(ctx, scopingUrl, desiredName, project.ProjectId, cancellationToken);
        var promoteSite = string.IsNullOrWhiteSpace(destination.SiteUrl) ? _options.SiteUrl : destination.SiteUrl!;

        var signature = ReconcileSignature.Compute(project, destination.PracticeLeaderEmail, destination.AdminEmails, destination.MembersPermissionLevel);
        await ApplyMetadataAsync(ctx, url, project, signature, promoteSite, cancellationToken);

        // Authoritative reset: the scoping grantees (the deal owner) give way to the delivery team.
        var addedByPromotion = await ApplyPermissionsAsync(ctx, url, project, destination, cancellationToken);
        await NotifyProjectAccessAddedAsync(ctx, project, destination, url, promoteSite, addedByPromotion, cancellationToken);

        // Deliberately NOT EnsureClientUploadsAsync — the scoping phase already created the folder and
        // minted the link (create-once). Re-running it would add a second folder and a duplicate link.
        var promoteUpload = await ReadUploadLinkAsync(ctx, url);
        return new DocumentSetResult(Created: false, url, Promoted: true,
            DataroomUrl: BuildAbsoluteUrl(promoteSite, url), ClientUploadUrl: promoteUpload);
    }

    /// <summary>
    /// Renames a document set folder in place via FileLeafRef — a rename, not a move: contents, item id, and
    /// sharing links all follow it. Returns the resulting server-relative URL (a rename changes it), or the
    /// original URL when the name is already correct or the rename fails. Fail-soft by design: a folder name
    /// is cosmetic and not worth aborting a promotion over.
    /// </summary>
    private async Task<string> TryRenameDocumentSetAsync(
        ClientContext ctx, string serverRelativeUrl, string desiredName, string discriminator, CancellationToken cancellationToken)
    {
        try
        {
            var folder = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl));
            ctx.Load(folder, f => f.Name, f => f.ParentFolder);
            var item = folder.ListItemAllFields;
            ctx.Load(item);
            await ctx.ExecuteQueryRetryAsync();

            var currentName = folder.Name;
            if (string.Equals(currentName, desiredName, StringComparison.Ordinal))
            {
                return serverRelativeUrl;
            }

            // A sibling may already hold the desired name (same customer prefix, different engagement).
            var uniqueName = await ResolveUniqueNameAsync(ctx, folder.ParentFolder, desiredName, discriminator, cancellationToken);
            item["FileLeafRef"] = uniqueName;
            item.Update();
            await ctx.ExecuteQueryRetryAsync();

            ctx.Load(item, i => i["FileRef"]);
            await ctx.ExecuteQueryRetryAsync();

            _logger.LogInformation("Renamed document set {OldName} to {NewName}.", currentName, uniqueName);
            var renamedUrl = item["FileRef"]?.ToString();
            return string.IsNullOrWhiteSpace(renamedUrl) ? serverRelativeUrl : renamedUrl!;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Failed to rename the document set at {Url} to {Name}; keeping the existing name.",
                serverRelativeUrl, desiredName);
            return serverRelativeUrl;
        }
    }

    private async Task ApplyMetadataAsync(
        ClientContext ctx, string serverRelativeUrl, AcumaticaProject project, string? signature, string siteUrl, CancellationToken cancellationToken)
    {
        var folder = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl));
        var item = folder.ListItemAllFields;
        try
        {
            await EnsurePrincipalInChargeColumnAsync(ctx, item.ParentList);
        }
        catch (Exception ex)
        {
            // Never block the room over the PIC column; it's retried on the next write.
            _logger.LogWarning("Couldn't ensure the Principal in Charge column: {Message}", ex.Message);
        }
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

        // PIC (People column). Resolved up front for the same single-Update reason; blank clears it.
        var hasPicColumn = !string.IsNullOrWhiteSpace(_options.PrincipalInChargeColumn)
            && item.FieldValues.ContainsKey(_options.PrincipalInChargeColumn);
        FieldUserValue? picValue = null;
        if (hasPicColumn && !string.IsNullOrWhiteSpace(project.PrincipalInChargeEmail))
        {
            var pic = await TryEnsureUserAsync(ctx, project.PrincipalInChargeEmail);
            picValue = pic is null ? null : new FieldUserValue { LookupId = pic.Id };
        }

        SetIfPresent(item, _options.ProjectIdColumn, project.ProjectId);
        SetIfPresent(item, _options.CustomerNameColumn, project.CustomerName);
        SetIfPresent(item, _options.ProjectNameColumn, project.ProjectName);
        SetIfPresent(item, _options.PracticeColumn, project.Practice);
        SetIfPresent(item, _options.StatusColumn, _options.ProjectStatusValue);

        // Deliberately does NOT touch the HubSpot deal-id / opportunity-number columns: those are written
        // by the scoping phase and are what let this workspace be recognised as already-promoted. Acumatica
        // only knows the PQCode value, which may be either of them — overwriting one from the other would
        // corrupt the link. Project Id + Status above are what mark the promotion.

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

        if (hasPicColumn)
        {
            item[_options.PrincipalInChargeColumn] = picValue;
        }

        // Stamp the reconcile signature so an unchanged project is skipped on later sweeps.
        SetIfPresent(item, _options.SignatureColumn, signature);

        // Keep the dataroom URL in the metadata column current. This is the value the reconcile sweep
        // compares against to notice a manual folder rename (and then re-write Acumatica's DATAURL).
        SetIfPresent(item, _options.DataroomUrlColumn, BuildAbsoluteUrl(siteUrl, serverRelativeUrl));

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
    private Task<IReadOnlyList<string>> ApplyPermissionsAsync(
        ClientContext ctx, string serverRelativeUrl, AcumaticaProject project, PracticeMappingEntry destination, CancellationToken cancellationToken)
    {
        // Everyone who should get the grant level: PM, practice leader, and the project team.
        var pmIdentity = !string.IsNullOrWhiteSpace(project.ProjectManagerEmail)
            ? project.ProjectManagerEmail
            : project.ProjectManager;
        var identities = new List<string?> { pmIdentity, project.PrincipalInChargeEmail, destination.PracticeLeaderEmail };
        identities.AddRange(project.TeamEmails);
        identities.AddRange(destination.AdminEmails);
        return ApplyPermissionsCoreAsync(ctx, serverRelativeUrl, identities, destination.MembersPermissionLevel,
            $"project {project.ProjectId}", cancellationToken);
    }

    /// <summary>
    /// Grants each of <paramref name="granteeIdentities"/> the configured level on the document set; when
    /// RestrictPermissions is on, breaks inheritance so only they and the site Owners have access.
    /// Fail-soft on individual unresolvable users; a hard permission-API failure is logged (surfaces via
    /// alerting) but does not roll back the created set.
    /// </summary>
    /// <summary>
    /// Applies the grant set and returns the identities that are NEWLY granted on this run (those not
    /// already assigned before it) — the basis for "you've been added" notifications. Returns an empty
    /// list when permissions are disabled or on any failure (so callers never notify on a partial state).
    /// </summary>
    private async Task<IReadOnlyList<string>> ApplyPermissionsCoreAsync(
        ClientContext ctx, string serverRelativeUrl, IReadOnlyList<string?> granteeIdentities, string? membersLevel,
        string logContext, CancellationToken cancellationToken)
    {
        if (!_options.SetProjectPermissions)
        {
            return Array.Empty<string>();
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
            RoleDefinition? membersRole = null;
            if (!string.IsNullOrWhiteSpace(membersLevel))
            {
                membersRole = web.RoleDefinitions.GetByName(membersLevel.Trim());
                ctx.Load(membersRole);
                ctx.Load(web.AssociatedMemberGroup);
            }
            // Who is assigned BEFORE this run, so we can tell which grantees are new.
            ctx.Load(item.RoleAssignments, r => r.Include(a => a.PrincipalId));
            await ctx.ExecuteQueryRetryAsync();

            var priorPrincipalIds = item.RoleAssignments.Select(a => a.PrincipalId).ToHashSet();

            // Resolve to distinct site users (fail-soft; external/unresolvable emails are skipped),
            // remembering each user's originating identity so we can report the ones newly added.
            var grantees = new List<User>();
            var identityByUserId = new Dictionary<int, string>();
            foreach (var identity in granteeIdentities)
            {
                var user = await TryEnsureUserAsync(ctx, identity);
                if (user is not null && !identityByUserId.ContainsKey(user.Id))
                {
                    grantees.Add(user);
                    identityByUserId[user.Id] = identity!.Trim();
                }
            }

            var newlyAdded = grantees
                .Where(u => !priorPrincipalIds.Contains(u.Id))
                .Select(u => identityByUserId[u.Id])
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (_options.RestrictPermissions)
            {
                // Clean slate, then re-grant. Site collection admins / the app identity retain access.
                // When client-upload links are enabled we must NOT clear sub-scopes, or a later reconcile
                // would wipe the child "Client Uploads" folder's upload link (which is create-once). The
                // doc set's own permissions are still fully reset either way.
                item.BreakRoleInheritance(copyRoleAssignments: false, clearSubscopes: !_options.CreateClientUploadLink);
                var ownerBinding = new RoleDefinitionBindingCollection(ctx);
                ownerBinding.Add(fullControl);
                item.RoleAssignments.Add(web.AssociatedOwnerGroup, ownerBinding);
            }

            foreach (var user in grantees)
            {
                var binding = new RoleDefinitionBindingCollection(ctx);
                binding.Add(grantRole);
                item.RoleAssignments.Add(user, binding);
            }

            // Practice-wide access: the site Members group (a group, so it's never in the "newly added" emails).
            if (membersRole is not null)
            {
                var membersBinding = new RoleDefinitionBindingCollection(ctx);
                membersBinding.Add(membersRole);
                item.RoleAssignments.Add(web.AssociatedMemberGroup, membersBinding);
            }

            await ctx.ExecuteQueryRetryAsync();
            _logger.LogInformation(
                "Set permissions on {Context}: {Count} user(s) at {Level} + Owners FullControl{Members} ({New} newly added).",
                logContext, grantees.Count, _options.PermissionLevel,
                membersRole is null ? "" : $" + Members {membersLevel}", newlyAdded.Count);
            return newlyAdded;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "Failed to set permissions on the document set for {Context}. The set exists with its current permissions; re-run to retry.",
                logContext);
            return Array.Empty<string>();
        }
    }

    /// <summary>
    /// Creates the "Client Uploads" subfolder inside a document set and mints an anonymous upload-only
    /// ("Request files") link for it via Graph, stamping the URL into <see cref="SharePointOptions.ClientUploadLinkColumn"/>.
    /// Fully fail-soft: a failure here is logged but does not fail document-set creation.
    /// </summary>
    private async Task<string?> EnsureClientUploadsAsync(
        ClientContext ctx, List list, string docSetServerRelativeUrl, string siteUrl, CancellationToken cancellationToken)
    {
        try
        {
            var folderName = SharePointNaming.SanitizeLeafName(_options.ClientUploadsFolderName);
            var docSetFolder = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(docSetServerRelativeUrl));
            var uploads = docSetFolder.Folders.Add(folderName);
            ctx.Load(uploads, f => f.ServerRelativeUrl);
            ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
            await ctx.ExecuteQueryRetryAsync();

            var link = await _uploadLinks.CreateUploadLinkAsync(
                siteUrl, list.RootFolder.ServerRelativeUrl, uploads.ServerRelativeUrl, cancellationToken);

            var col = _options.ClientUploadLinkColumn;
            if (!string.IsNullOrWhiteSpace(link) && !string.IsNullOrWhiteSpace(col))
            {
                // A plain Text column (not a Hyperlink): the value is meant to be copied and sent to the
                // client, so it shows as a copyable URL string rather than a click target that navigates
                // the (signed-in) user to the folder instead of the upload page.
                await EnsureUploadLinkColumnAsync(ctx, list, col);

                // Stamp the link on the document set item (the folder the metadata lives on), not the subfolder.
                var docItem = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(docSetServerRelativeUrl)).ListItemAllFields;
                ctx.Load(docItem);
                await ctx.ExecuteQueryRetryAsync();
                SetIfPresent(docItem, col, link);
                docItem.Update();
                await ctx.ExecuteQueryRetryAsync();
            }

            return link;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Failed to create the Client Uploads folder/link for document set at {Url}; the set exists without it.",
                docSetServerRelativeUrl);
            return null;
        }
    }

    private async Task ApplyScopingMetadataAsync(
        ClientContext ctx, string serverRelativeUrl, ScopingWorkspace ws, CancellationToken cancellationToken)
    {
        var folder = ctx.Web.GetFolderByServerRelativeUrl(ToServerRelative(serverRelativeUrl));
        var item = folder.ListItemAllFields;
        ctx.Load(item);
        await ctx.ExecuteQueryRetryAsync();

        // Resolve the PM People field first (own round-trip), so field writes stay in one Update. Prefer the
        // deal's actual project manager; fall back to the deal owner when no PM is set on the deal.
        var pmIdentity = !string.IsNullOrWhiteSpace(ws.ProjectManagerEmail) ? ws.ProjectManagerEmail : ws.OwnerEmail;
        FieldUserValue? pmValue = null;
        if (_options.ProjectManagerIsPersonColumn && !string.IsNullOrWhiteSpace(pmIdentity))
        {
            pmValue = await ResolvePersonAsync(ctx, _options.ProjectManagerColumn, item, pmIdentity, null);
        }

        SetIfPresent(item, _options.CustomerNameColumn, ws.CustomerName);
        SetIfPresent(item, _options.ProjectNameColumn, ws.ProjectName);
        SetIfPresent(item, _options.PracticeColumn, ws.Practice);
        SetIfPresent(item, _options.HubSpotDealIdColumn, ws.DealId);
        SetIfPresent(item, _options.OpportunityIdColumn, ws.OpportunityId);
        SetIfPresent(item, _options.StatusColumn, _options.ScopingStatusValue);

        if (_options.ProjectManagerIsPersonColumn && pmValue is not null && item.FieldValues.ContainsKey(_options.ProjectManagerColumn))
        {
            item[_options.ProjectManagerColumn] = pmValue;
        }

        item.Update();
        await ctx.ExecuteQueryRetryAsync();
    }

    private Task<IReadOnlyList<string>> ApplyScopingPermissionsAsync(
        ClientContext ctx, string serverRelativeUrl, ScopingWorkspace ws, PracticeMappingEntry destination, CancellationToken cancellationToken)
    {
        // Scoping access: the deal owner + project manager + originators A/B + the practice leader + admins.
        var identities = new List<string?>
        {
            ws.OwnerEmail, ws.ProjectManagerEmail, ws.OriginatorAEmail, ws.OriginatorBEmail,
            destination.PracticeLeaderEmail,
        };
        identities.AddRange(destination.AdminEmails);
        return ApplyPermissionsCoreAsync(ctx, serverRelativeUrl, identities, destination.MembersPermissionLevel,
            $"deal {ws.DealId}", cancellationToken);
    }

    // Libraries already checked for the PIC column in this process (one lookup per library, not per room).
    // The service is a singleton shared by concurrently running functions, hence the concurrent set.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<Guid, bool> _picColumnChecked = new();

    /// <summary>Ensures the Principal in Charge People column exists on the library (visible in the default view).</summary>
    private async Task EnsurePrincipalInChargeColumnAsync(ClientContext ctx, List list)
    {
        var col = _options.PrincipalInChargeColumn;
        if (string.IsNullOrWhiteSpace(col))
        {
            return;
        }

        ctx.Load(list, l => l.Id);
        await ctx.ExecuteQueryRetryAsync();
        if (_picColumnChecked.ContainsKey(list.Id))
        {
            return;
        }

        ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName));
        await ctx.ExecuteQueryRetryAsync();
        if (!list.Fields.Any(f => f.InternalName == col))
        {
            // Created with the internal name, then renamed so people see "Principal in Charge".
            var field = list.Fields.AddFieldAsXml(
                $"<Field Type='User' Name='{col}' StaticName='{col}' DisplayName='{col}' UserSelectionMode='PeopleOnly' Group='ProjectSync'/>",
                addToDefaultView: true, options: AddFieldOptions.AddFieldInternalNameHint);
            field.Title = "Principal in Charge";
            field.Update();
            await ctx.ExecuteQueryRetryAsync();
            _logger.LogInformation("Created People column '{Column}' (Principal in Charge).", col);
        }

        _picColumnChecked.TryAdd(list.Id, true);
    }

    /// <summary>Ensures a visible text column exists on the library, creating it if missing.</summary>
    private async Task EnsureTextColumnAsync(ClientContext ctx, List list, string col)
    {
        if (string.IsNullOrWhiteSpace(col))
        {
            return;
        }

        ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName));
        await ctx.ExecuteQueryRetryAsync();
        if (list.Fields.Any(f => f.InternalName == col))
        {
            return;
        }

        list.Fields.AddFieldAsXml(
            $"<Field Type='Text' Name='{col}' StaticName='{col}' DisplayName='{col}' Group='ProjectSync'/>",
            addToDefaultView: true, options: AddFieldOptions.AddFieldInternalNameHint);
        await ctx.ExecuteQueryRetryAsync();
        _logger.LogInformation("Created column '{Column}'.", col);
    }

    /// <summary>
    /// Finds the document set (FSObjType=1) whose <paramref name="column"/> equals <paramref name="value"/>,
    /// returning its URL together with its Project Id column — a blank Project Id means the workspace is
    /// still in the scoping phase, a populated one means it has already been promoted.
    /// </summary>
    private async Task<TrackedDocSet?> FindByColumnAsync(
        ClientContext ctx, List list, string? parentFolder, string column, string value, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(column) || string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var pidCol = _options.ProjectIdColumn;
        var items = await FindRoomsByColumnAsync(ctx, list, parentFolder, column, value, 2, new[] { pidCol }, cancellationToken);
        if (items.Count == 0)
        {
            return null;
        }

        if (items.Count > 1)
        {
            // Two document sets share this value. HubSpot does contain the occasional duplicate opportunity
            // number, so this is reachable via the PQCode → OpportunityId match. Binding to whichever row
            // came back first would silently attach an engagement to another one's folder — wrong metadata
            // and wrong permissions on a client folder — so refuse and let the caller create its own. The
            // warning is the signal to fix the source data.
            _logger.LogWarning(
                "Column {Column} value {Value} matches {Count}+ document sets; refusing to guess which. Resolve the duplicate in the source system.",
                column, value, items.Count);
            return null;
        }

        var url = items[0]["FileRef"]?.ToString();
        if (string.IsNullOrWhiteSpace(url))
        {
            return null;
        }

        var projectId = items[0].FieldValues.TryGetValue(pidCol, out var pid) ? pid?.ToString()?.Trim() : null;
        return new TrackedDocSet(url!, projectId);
    }

    /// <summary>A document set located by lookup: where it lives, and the project id it carries (if any).</summary>
    private sealed record TrackedDocSet(string Url, string? ProjectId);

    /// <summary>
    /// Ensures the client-upload-link column exists as a plain single-line Text field (so the URL is a
    /// copyable string, not a click target). Migrates an older Hyperlink/URL column to Text if found.
    /// </summary>
    private async Task EnsureUploadLinkColumnAsync(ClientContext ctx, List list, string col)
    {
        try
        {
            ctx.Load(list.Fields, fs => fs.Include(f => f.InternalName, f => f.FieldTypeKind));
            await ctx.ExecuteQueryRetryAsync();

            var existing = list.Fields.FirstOrDefault(f => f.InternalName == col);
            if (existing is not null)
            {
                if (existing.FieldTypeKind == FieldType.Text)
                {
                    return;
                }

                // Migrate the earlier Hyperlink design to Text so the value copies as a plain URL string.
                _logger.LogInformation("Migrating upload-link column '{Column}' from {Type} to Text.", col, existing.FieldTypeKind);
                existing.DeleteObject();
                await ctx.ExecuteQueryRetryAsync();
            }

            list.Fields.AddFieldAsXml(
                $"<Field Type='Text' Name='{col}' StaticName='{col}' DisplayName='{col}' Group='ProjectSync'/>",
                addToDefaultView: true, options: AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryRetryAsync();
            _logger.LogInformation("Ensured client-upload-link column '{Column}' (Text).", col);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not ensure client-upload-link column '{Column}'; the link won't be stamped.", col);
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
