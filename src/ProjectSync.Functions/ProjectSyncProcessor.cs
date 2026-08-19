using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

namespace ProjectSync;

/// <summary>
/// Core workflow: read the last-run watermark, pull newly-created projects from Acumatica,
/// create/ensure a SharePoint document set for each, then advance the watermark.
/// </summary>
public sealed class ProjectSyncProcessor
{
    private readonly IAcumaticaClient _acumatica;
    private readonly ISharePointDocumentSetService _sharePoint;
    private readonly ILastRunStore _lastRunStore;
    private readonly StateOptions _stateOptions;
    private readonly TimeProvider _timeProvider;
    private readonly ILogger<ProjectSyncProcessor> _logger;
    private readonly HashSet<string> _includedPractices;
    private readonly HashSet<string> _excludedProjectIds;

    public ProjectSyncProcessor(
        IAcumaticaClient acumatica,
        ISharePointDocumentSetService sharePoint,
        ILastRunStore lastRunStore,
        IOptions<StateOptions> stateOptions,
        IOptions<AcumaticaOptions> acumaticaOptions,
        TimeProvider timeProvider,
        ILogger<ProjectSyncProcessor> logger)
    {
        _acumatica = acumatica;
        _sharePoint = sharePoint;
        _lastRunStore = lastRunStore;
        _stateOptions = stateOptions.Value;
        _timeProvider = timeProvider;
        _logger = logger;
        _includedPractices = new HashSet<string>(
            acumaticaOptions.Value.IncludedPractices.Select(p => p.Trim()),
            StringComparer.OrdinalIgnoreCase);
        _excludedProjectIds = new HashSet<string>(
            acumaticaOptions.Value.ExcludedProjectIds.Select(p => p.Trim()),
            StringComparer.OrdinalIgnoreCase);
    }

    public Task<ProjectSyncResult> RunAsync(CancellationToken cancellationToken)
        => RunAsync(new RunOptions(), cancellationToken);

    public async Task<ProjectSyncResult> RunAsync(RunOptions runOptions, CancellationToken cancellationToken)
    {
        var dryRun = runOptions.DryRun;
        var lastRun = await _lastRunStore.GetLastRunAsync(cancellationToken);
        var nowUtc = _timeProvider.GetUtcNow();

        // Determine the query lower bound. A caller may override it (dry-run exploration); otherwise
        // on first run look back a bounded window, and on later runs apply a small overlap to guard
        // against commit latency / clock skew.
        var queryFrom = runOptions.OverrideSince
            ?? (lastRun is null
                ? nowUtc.AddHours(-_stateOptions.FirstRunLookbackHours)
                : lastRun.Value.AddMinutes(-_stateOptions.OverlapMinutes));

        _logger.LogInformation("ProjectSync starting{DryRun}. LastRun={LastRun:o}, querying created > {From:o}",
            dryRun ? " (DRY RUN)" : string.Empty, lastRun, queryFrom);

        IReadOnlyList<AcumaticaProject> projects = await _acumatica.GetProjectsCreatedAfterAsync(queryFrom, cancellationToken);

        // Targeted single-project run: filter to just that project and never persist the watermark.
        var isTargeted = !string.IsNullOrWhiteSpace(runOptions.OnlyProjectId);
        if (isTargeted)
        {
            projects = projects
                .Where(p => string.Equals(p.ProjectId.Trim(), runOptions.OnlyProjectId!.Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();
            _logger.LogInformation("Targeted run for project {ProjectId}: {Count} match(es).", runOptions.OnlyProjectId, projects.Count);
        }

        var persistWatermark = !dryRun && !isTargeted;

        if (projects.Count == 0)
        {
            _logger.LogInformation("No new projects found.");
            // Advance watermark so first-run lookback doesn't repeat every cycle (skip on dry/targeted run).
            if (lastRun is null && persistWatermark)
            {
                await _lastRunStore.SetLastRunAsync(nowUtc, cancellationToken);
            }

            return new ProjectSyncResult { DryRun = dryRun, Found = 0, Watermark = lastRun ?? nowUtc, Plan = dryRun ? Array.Empty<PlannedDocumentSet>() : null };
        }

        var watermark = lastRun ?? queryFrom;
        var created = 0;
        var promoted = 0;
        var updated = 0;
        var planned = 0;
        var skipped = 0;
        var hadFailure = false;
        var plan = dryRun ? new List<PlannedDocumentSet>() : null;

        foreach (var project in projects)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // Excluded project IDs (internal/system monikers, e.g. "X"). Hard guard independent of
            // practice. Skipped projects are treated as handled and still advance the watermark.
            if (_excludedProjectIds.Contains(project.ProjectId.Trim()))
            {
                skipped++;
                _logger.LogDebug("Skipping project {ProjectId}: on the excluded-project-id list.", project.ProjectId);
                if (project.CreatedDateTime is { } excludedAt && excludedAt > watermark)
                {
                    watermark = excludedAt;
                }
                continue;
            }

            // Practice allow-list. Skipped projects are treated as successfully handled (no-op) and
            // still advance the watermark so they are not re-read on every cycle.
            if (_includedPractices.Count > 0 &&
                !_includedPractices.Contains((project.Practice ?? string.Empty).Trim()))
            {
                skipped++;
                _logger.LogDebug("Skipping project {ProjectId}: practice '{Practice}' not in the allow-list.",
                    project.ProjectId, project.Practice ?? "<none>");
                if (project.CreatedDateTime is { } skippedAt && skippedAt > watermark)
                {
                    watermark = skippedAt;
                }
                continue;
            }

            try
            {
                if (dryRun)
                {
                    // Resolve the destination from config only — no SharePoint connection.
                    var p = _sharePoint.PlanDocumentSet(project);
                    plan!.Add(new PlannedDocumentSet
                    {
                        ProjectId = project.ProjectId,
                        ProjectName = project.ProjectName,
                        CustomerName = project.CustomerName,
                        ProjectManager = project.ProjectManager,
                        ProjectManagerEmail = project.ProjectManagerEmail,
                        Practice = project.Practice,
                        HubSpotLink = project.HubSpotLink,
                        CreatedDateTime = project.CreatedDateTime,
                        TargetSiteUrl = p.SiteUrl,
                        TargetLibrary = p.Library,
                        TargetFolder = p.ParentFolder,
                        DocumentSetName = p.SetName,
                    });
                    planned++;
                }
                else
                {
                    // Enrich with the project team (for permissions) just before writing.
                    var teamEmails = await _acumatica.GetTeamEmailsAsync(project.ProjectId, cancellationToken);
                    var enriched = project with { TeamEmails = teamEmails };
                    var result = await _sharePoint.EnsureProjectDocumentSetAsync(enriched, cancellationToken);
                    if (result.Created)
                    {
                        created++;
                    }
                    else if (result.Promoted)
                    {
                        promoted++;
                    }
                    else
                    {
                        updated++;
                    }

                    // Once, on first creation (or promotion): write the workspace URLs back to the
                    // Acumatica project attributes. Fail-soft — never let this break the cycle.
                    if (result.Created || result.Promoted)
                    {
                        try
                        {
                            await _acumatica.WriteProjectUrlsAsync(
                                project.ProjectId, result.DataroomUrl, result.ClientUploadUrl, cancellationToken);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "URL write-back to Acumatica failed for project {ProjectId}; continuing.", project.ProjectId);
                        }
                    }
                }

                // Advance the watermark only for successfully processed records, oldest-first, so a
                // failure leaves the watermark just before the offending record for retry next cycle.
                if (project.CreatedDateTime is { } createdAt && createdAt > watermark)
                {
                    watermark = createdAt;
                }
            }
            catch (Exception ex)
            {
                hadFailure = true;
                _logger.LogError(ex,
                    "Failed to {Action} document set for project {ProjectId}. Halting this cycle; watermark held at {Watermark:o} for retry.",
                    dryRun ? "plan" : "create", project.ProjectId, watermark);
                break;
            }
        }

        // A dry run or a targeted single-project run never mutates the watermark.
        if (persistWatermark)
        {
            await _lastRunStore.SetLastRunAsync(watermark, cancellationToken);
        }

        _logger.LogInformation(
            "ProjectSync complete{DryRun}. Found={Found}, Created={Created}, Promoted={Promoted}, UpdatedExisting={Updated}, Planned={Planned}, Skipped={Skipped}, watermark={Watermark:o}",
            dryRun ? " (DRY RUN)" : string.Empty, projects.Count, created, promoted, updated, planned, skipped, watermark);

        return new ProjectSyncResult
        {
            DryRun = dryRun,
            Found = projects.Count,
            Created = created,
            Updated = updated,
            Promoted = promoted,
            Planned = planned,
            Skipped = skipped,
            HadFailure = hadFailure,
            Watermark = watermark,
            Plan = plan,
        };
    }

    private const string ReconcileTeamWatermark = "reconcile-team";

    /// <summary>
    /// Incremental reconcile: uses the team GI's modified date to touch only projects whose team
    /// changed since the last pass. Short-circuits to zero SharePoint work when nothing changed.
    /// </summary>
    public async Task<ReconcileResult> ReconcileIncrementalAsync(CancellationToken cancellationToken)
    {
        var teamRows = await _acumatica.GetTeamRowsAsync(cancellationToken);
        var watermark = await _lastRunStore.GetWatermarkAsync(ReconcileTeamWatermark, cancellationToken);

        var changed = teamRows
            .Where(r => r.ModifiedOn is { } m && (watermark is null || m > watermark))
            .Select(r => r.ProjectId.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (changed.Count == 0)
        {
            _logger.LogInformation("Reconcile (incremental): no team changes since {Watermark:o}.", watermark);
            // Establish the baseline on first run so we don't reprocess history every cycle.
            if (watermark is null)
            {
                await _lastRunStore.SetWatermarkAsync(ReconcileTeamWatermark, MaxModified(teamRows), cancellationToken);
            }
            return new ReconcileResult();
        }

        var desired = await BuildDesiredProjectsAsync(teamRows, cancellationToken);
        var result = await _sharePoint.ReconcileAsync(desired, changed, cancellationToken);
        await ResyncRenamedUrlsAsync(result, cancellationToken);
        await _lastRunStore.SetWatermarkAsync(ReconcileTeamWatermark, MaxModified(teamRows), cancellationToken);

        _logger.LogInformation("Reconcile (incremental): {Changed} team-changed project(s); updated {Updated}, unchanged {Unchanged}.",
            changed.Count, result.Updated, result.Unchanged);
        return result;
    }

    /// <summary>Daily full reconcile of every tracked document set (signature-gated; catches removals + metadata).</summary>
    public async Task<ReconcileResult> ReconcileFullAsync(CancellationToken cancellationToken)
    {
        var teamRows = await _acumatica.GetTeamRowsAsync(cancellationToken);
        var desired = await BuildDesiredProjectsAsync(teamRows, cancellationToken);
        var result = await _sharePoint.ReconcileAsync(desired, onlyProjectIds: null, cancellationToken);
        await ResyncRenamedUrlsAsync(result, cancellationToken);

        // A full sweep covers everything up to now — advance the incremental cursor too.
        await _lastRunStore.SetWatermarkAsync(ReconcileTeamWatermark, MaxModified(teamRows), cancellationToken);

        _logger.LogInformation("Reconcile (full): considered {Considered}, updated {Updated}, unchanged {Unchanged}, notTracked {NotTracked}.",
            result.Considered, result.Updated, result.Unchanged, result.NotTracked);
        return result;
    }

    /// <summary>
    /// Re-writes Acumatica's DATAURL for any workspace whose SharePoint folder was renamed since the last
    /// sweep (detected by the reconcile). Fail-soft per project — a write-back failure never breaks the pass.
    /// </summary>
    private async Task ResyncRenamedUrlsAsync(ReconcileResult result, CancellationToken cancellationToken)
    {
        foreach (var r in result.UrlResyncs)
        {
            try
            {
                await _acumatica.WriteProjectUrlsAsync(r.ProjectId, r.DataroomUrl, clientUrl: null, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "DATAURL resync to Acumatica failed for project {ProjectId}; continuing.", r.ProjectId);
            }
        }
    }

    private DateTimeOffset MaxModified(IReadOnlyList<TeamMemberRow> teamRows)
        => teamRows.Where(r => r.ModifiedOn is not null)
            .Select(r => r.ModifiedOn!.Value)
            .DefaultIfEmpty(_timeProvider.GetUtcNow())
            .Max();

    /// <summary>Builds the desired projects (included practices, enriched with their current team) from the GIs.</summary>
    private async Task<IReadOnlyList<AcumaticaProject>> BuildDesiredProjectsAsync(
        IReadOnlyList<TeamMemberRow> teamRows, CancellationToken cancellationToken)
    {
        var teamByProject = teamRows
            .GroupBy(r => r.ProjectId.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => (IReadOnlyList<string>)g.Select(r => r.Email.Trim())
                    .Where(e => !string.IsNullOrWhiteSpace(e))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                StringComparer.OrdinalIgnoreCase);

        // All projects (no created bound), then the same include/exclude filters as the create path.
        var all = await _acumatica.GetProjectsCreatedAfterAsync(
            new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero), cancellationToken);

        var desired = new List<AcumaticaProject>();
        foreach (var p in all)
        {
            if (_excludedProjectIds.Contains(p.ProjectId.Trim()))
            {
                continue;
            }

            if (_includedPractices.Count > 0 && !_includedPractices.Contains((p.Practice ?? string.Empty).Trim()))
            {
                continue;
            }

            var team = teamByProject.TryGetValue(p.ProjectId.Trim(), out var t) ? t : Array.Empty<string>();
            desired.Add(p with { TeamEmails = team });
        }

        return desired;
    }
}

/// <summary>Options controlling a single <see cref="ProjectSyncProcessor.RunAsync(RunOptions, CancellationToken)"/> invocation.</summary>
public sealed record RunOptions
{
    /// <summary>When true, resolve destinations and report a plan but create nothing and do not move the watermark.</summary>
    public bool DryRun { get; init; }

    /// <summary>Optional query lower-bound override (used by dry-run to widen the window). Null = normal watermark logic.</summary>
    public DateTimeOffset? OverrideSince { get; init; }

    /// <summary>
    /// When set, process only the project with this exact id (targeted reprocess/backfill-one). The
    /// watermark is NOT advanced for a targeted run, so it doesn't disturb the normal moving-forward baseline.
    /// </summary>
    public string? OnlyProjectId { get; init; }
}
