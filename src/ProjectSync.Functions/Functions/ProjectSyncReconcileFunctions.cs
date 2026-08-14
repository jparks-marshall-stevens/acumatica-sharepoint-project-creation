using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace ProjectSync.Functions;

/// <summary>
/// Keeps already-created document sets in sync with Acumatica:
///  - incremental (frequent): only projects whose team changed since the last pass (team GI ModifiedOn);
///  - full (daily): every tracked set, to catch team removals + PM/Description changes.
/// Both are signature-gated, so unchanged sets cost no writes.
/// </summary>
public sealed class ProjectSyncReconcileFunctions
{
    private readonly ProjectSyncProcessor _processor;
    private readonly ILogger<ProjectSyncReconcileFunctions> _logger;

    public ProjectSyncReconcileFunctions(ProjectSyncProcessor processor, ILogger<ProjectSyncReconcileFunctions> logger)
    {
        _processor = processor;
        _logger = logger;
    }

    [Function("ProjectSyncReconcileIncremental")]
    public async Task Incremental(
        [TimerTrigger("%ProjectSyncReconcileSchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        try
        {
            var result = await _processor.ReconcileIncrementalAsync(cancellationToken);
            _logger.LogInformation("Reconcile (incremental) done: updated {Updated}, unchanged {Unchanged}.",
                result.Updated, result.Unchanged);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Incremental reconcile failed.");
            throw;
        }
    }

    [Function("ProjectSyncReconcileFull")]
    public async Task Full(
        [TimerTrigger("%ProjectSyncFullReconcileSchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        try
        {
            var result = await _processor.ReconcileFullAsync(cancellationToken);
            _logger.LogInformation("Reconcile (full) done: considered {Considered}, updated {Updated}.",
                result.Considered, result.Updated);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Full reconcile failed.");
            throw;
        }
    }
}
