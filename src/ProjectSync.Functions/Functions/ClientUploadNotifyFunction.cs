using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace ProjectSync.Functions;

/// <summary>
/// Timer that notifies workspace members when the client uploads new files. Schedule is overridable via
/// the "ClientUploadNotifySchedule" app setting (NCRONTAB).
/// </summary>
public sealed class ClientUploadNotifyFunction
{
    private readonly ClientUploadProcessor _processor;
    private readonly ILogger<ClientUploadNotifyFunction> _logger;

    public ClientUploadNotifyFunction(ClientUploadProcessor processor, ILogger<ClientUploadNotifyFunction> logger)
    {
        _processor = processor;
        _logger = logger;
    }

    [Function("ClientUploadNotify")]
    public async Task Run(
        [TimerTrigger("%ClientUploadNotifySchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("ClientUploadNotify fired at {Time:o} (past due: {PastDue}).",
            DateTimeOffset.UtcNow, timer.IsPastDue);

        try
        {
            var result = await _processor.RunAsync(cancellationToken);
            _logger.LogInformation("ClientUploadNotify done. NewFiles={Files} Workspaces={Workspaces} Notified={Notified}.",
                result.NewFiles, result.WorkspacesWithNewFiles, result.Notified);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Client-upload notify cycle failed.");
            throw;
        }
    }
}
