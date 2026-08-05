using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace ProjectSync.Functions;

public sealed class ProjectSyncTimerFunction
{
    private readonly ProjectSyncProcessor _processor;
    private readonly ILogger<ProjectSyncTimerFunction> _logger;

    public ProjectSyncTimerFunction(ProjectSyncProcessor processor, ILogger<ProjectSyncTimerFunction> logger)
    {
        _processor = processor;
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
        _logger.LogInformation("ProjectSyncTimer fired at {Time:o} (past due: {PastDue}).",
            DateTimeOffset.UtcNow, timer.IsPastDue);

        try
        {
            var result = await _processor.RunAsync(cancellationToken);
            _logger.LogInformation("ProjectSyncTimer done. Found={Found} Created={Created} Updated={Updated}.",
                result.Found, result.Created, result.Updated);
        }
        catch (Exception ex)
        {
            // Let the exception surface so the invocation is recorded as failed and retried on schedule.
            _logger.LogError(ex, "ProjectSync cycle failed.");
            throw;
        }
    }
}
