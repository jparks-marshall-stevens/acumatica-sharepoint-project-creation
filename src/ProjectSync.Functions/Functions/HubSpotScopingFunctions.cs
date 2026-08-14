using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using ProjectSync.HubSpot;

namespace ProjectSync.Functions;

/// <summary>
/// Polls HubSpot for deals in the scoping phase (practice in scope, not yet Won/Lost) and creates/updates
/// a SharePoint scoping workspace for each, keyed on the HubSpot deal id. Watermark-driven, so each run
/// processes only deals modified since the last poll.
/// </summary>
public sealed class HubSpotScopingFunctions
{
    private readonly HubSpotScopingProcessor _processor;
    private readonly ILogger<HubSpotScopingFunctions> _logger;

    public HubSpotScopingFunctions(HubSpotScopingProcessor processor, ILogger<HubSpotScopingFunctions> logger)
    {
        _processor = processor;
        _logger = logger;
    }

    [Function("HubSpotScopingPoll")]
    public async Task Poll(
        [TimerTrigger("%HubSpotScopingSchedule%")] TimerInfo timer,
        CancellationToken cancellationToken)
    {
        try
        {
            var result = await _processor.RunAsync(dryRun: false, cancellationToken);
            _logger.LogInformation(
                "HubSpot scoping poll done: in-scope {InScope}, created {Created}, updated {Updated}.",
                result.InScope, result.Created, result.Updated);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "HubSpot scoping poll failed.");
            throw;
        }
    }
}
