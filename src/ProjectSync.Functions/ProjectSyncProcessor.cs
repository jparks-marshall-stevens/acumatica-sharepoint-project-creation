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

    public ProjectSyncProcessor(
        IAcumaticaClient acumatica,
        ISharePointDocumentSetService sharePoint,
        ILastRunStore lastRunStore,
        IOptions<StateOptions> stateOptions,
        TimeProvider timeProvider,
        ILogger<ProjectSyncProcessor> logger)
    {
        _acumatica = acumatica;
        _sharePoint = sharePoint;
        _lastRunStore = lastRunStore;
        _stateOptions = stateOptions.Value;
        _timeProvider = timeProvider;
        _logger = logger;
    }

    public async Task<ProjectSyncResult> RunAsync(CancellationToken cancellationToken)
    {
        var lastRun = await _lastRunStore.GetLastRunAsync(cancellationToken);
        var nowUtc = _timeProvider.GetUtcNow();

        // Determine the query lower bound. On first run, look back a bounded window rather than
        // importing all history. Apply a small overlap to guard against commit latency / clock skew.
        var queryFrom = lastRun is null
            ? nowUtc.AddHours(-_stateOptions.FirstRunLookbackHours)
            : lastRun.Value.AddMinutes(-_stateOptions.OverlapMinutes);

        _logger.LogInformation("ProjectSync starting. LastRun={LastRun:o}, querying created > {From:o}",
            lastRun, queryFrom);

        var projects = await _acumatica.GetProjectsCreatedAfterAsync(queryFrom, cancellationToken);
        if (projects.Count == 0)
        {
            _logger.LogInformation("No new projects found.");
            // Advance watermark so first-run lookback doesn't repeat every cycle.
            if (lastRun is null)
            {
                await _lastRunStore.SetLastRunAsync(nowUtc, cancellationToken);
            }

            return new ProjectSyncResult { Found = 0, Watermark = lastRun ?? nowUtc };
        }

        var watermark = lastRun ?? queryFrom;
        var created = 0;
        var updated = 0;
        var hadFailure = false;

        foreach (var project in projects)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
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
                    "Failed to create document set for project {ProjectId}. Halting this cycle; watermark held at {Watermark:o} for retry.",
                    project.ProjectId, watermark);
                break;
            }
        }

        await _lastRunStore.SetLastRunAsync(watermark, cancellationToken);
        _logger.LogInformation(
            "ProjectSync complete. Found={Found}, Created={Created}, UpdatedExisting={Updated}, watermark={Watermark:o}",
            projects.Count, created, updated, watermark);

        return new ProjectSyncResult
        {
            Found = projects.Count,
            Created = created,
            Updated = updated,
            HadFailure = hadFailure,
            Watermark = watermark,
        };
    }
}
