using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Acumatica;
using ProjectSync.Options;

namespace ProjectSync.Functions;

public sealed class ProjectSyncTimerFunction
{
    private readonly ProjectSyncProcessor _processor;
    private readonly AcumaticaOptions _acumatica;
    private readonly TimeProvider _time;
    private readonly ILogger<ProjectSyncTimerFunction> _logger;

    public ProjectSyncTimerFunction(ProjectSyncProcessor processor, IOptions<AcumaticaOptions> acumatica,
        TimeProvider time, ILogger<ProjectSyncTimerFunction> logger)
    {
        _processor = processor;
        _acumatica = acumatica.Value;
        _time = time;
        _logger = logger;
    }

    /// <summary>
    /// Runs every 15 minutes. Schedule is overridable via the "ProjectSyncSchedule" app setting
    /// (NCRONTAB). RunOnStartup is false so deployments don't trigger an unexpected sync.
    /// </summary>
    [Function("ProjectSyncTimer")]
    public async Task Run(
        [TimerTrigger("%ProjectSyncSchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        var startedUtc = _time.GetUtcNow();
        _logger.LogInformation("ProjectSyncTimer fired at {Time:o} (past due: {PastDue}).",
            startedUtc, timer.IsPastDue);

        try
        {
            var result = await _processor.RunAsync(cancellationToken);
            _logger.LogInformation("ProjectSyncTimer done. Found={Found} Created={Created} Updated={Updated}.",
                result.Found, result.Created, result.Updated);
        }
        catch (Exception ex) when (AcumaticaSlowWindow.IsExpectedTimeout(ex, startedUtc, _acumatica, cancellationToken))
        {
            // Known daily Acumatica stall: skip quietly (no exception object, so nothing reaches the
            // exceptions table / error alert). The watermark didn't move, so the next run catches up.
            _logger.LogWarning("ProjectSync cycle skipped: Acumatica timed out during its daily slow window ({Message}).",
                ex.Message);
        }
        catch (Exception ex)
        {
            // Let the exception surface so the invocation is recorded as failed and retried on schedule.
            _logger.LogError(ex, "ProjectSync cycle failed.");
            throw;
        }
    }
}
