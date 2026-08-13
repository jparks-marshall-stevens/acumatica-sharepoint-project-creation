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
                    var result = await _sharePoint.EnsureProjectDocumentSetAsync(project, cancellationToken);
                    if (result.Created)
                    {
                        created++;
                    }
                    else
                    {
                        updated++;
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
            "ProjectSync complete{DryRun}. Found={Found}, Created={Created}, UpdatedExisting={Updated}, Planned={Planned}, Skipped={Skipped}, watermark={Watermark:o}",
            dryRun ? " (DRY RUN)" : string.Empty, projects.Count, created, updated, planned, skipped, watermark);

        return new ProjectSyncResult
        {
            DryRun = dryRun,
            Found = projects.Count,
            Created = created,
            Updated = updated,
            Planned = planned,
            Skipped = skipped,
            HadFailure = hadFailure,
            Watermark = watermark,
            Plan = plan,
        };
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
