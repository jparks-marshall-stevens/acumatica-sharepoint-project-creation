using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Acumatica;
using ProjectSync.Options;

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
    private readonly AcumaticaOptions _acumatica;
    private readonly TimeProvider _time;
    private readonly ILogger<ProjectSyncReconcileFunctions> _logger;

    public ProjectSyncReconcileFunctions(ProjectSyncProcessor processor, IOptions<AcumaticaOptions> acumatica,
        TimeProvider time, ILogger<ProjectSyncReconcileFunctions> logger)
    {
        _processor = processor;
        _acumatica = acumatica.Value;
        _time = time;
        _logger = logger;
    }

    [Function("ProjectSyncReconcileIncremental")]
    public async Task Incremental(
        [TimerTrigger("%ProjectSyncReconcileSchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        var startedUtc = _time.GetUtcNow();
        try
        {
            var result = await _processor.ReconcileIncrementalAsync(cancellationToken);
            _logger.LogInformation("Reconcile (incremental) done: updated {Updated}, unchanged {Unchanged}.",
                result.Updated, result.Unchanged);
        }
        catch (Exception ex) when (AcumaticaSlowWindow.IsExpectedTimeout(ex, startedUtc, _acumatica, cancellationToken))
        {
            _logger.LogWarning("Incremental reconcile skipped: Acumatica timed out during its daily slow window ({Message}).",
                ex.Message);
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
        var startedUtc = _time.GetUtcNow();
        try
        {
            var result = await _processor.ReconcileFullAsync(cancellationToken);
            _logger.LogInformation("Reconcile (full) done: considered {Considered}, updated {Updated}.",
                result.Considered, result.Updated);
        }
        catch (Exception ex) when (AcumaticaSlowWindow.IsExpectedTimeout(ex, startedUtc, _acumatica, cancellationToken))
        {
            _logger.LogWarning("Full reconcile skipped: Acumatica timed out during its daily slow window ({Message}).",
                ex.Message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Full reconcile failed.");
            throw;
        }
    }
}
