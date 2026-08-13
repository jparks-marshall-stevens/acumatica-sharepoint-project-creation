using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace ProjectSync.Functions;

/// <summary>
/// Manual "run now" trigger for testing/backfills. Runs the exact same cycle as the timer
/// and returns a JSON summary. Protected by a function key (AuthorizationLevel.Function).
/// </summary>
public sealed class ProjectSyncHttpFunction
{
    private readonly ProjectSyncProcessor _processor;
    private readonly ILogger<ProjectSyncHttpFunction> _logger;

    public ProjectSyncHttpFunction(ProjectSyncProcessor processor, ILogger<ProjectSyncHttpFunction> logger)
    {
        _processor = processor;
        _logger = logger;
    }

    [Function("ProjectSyncRunNow")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "sync/run")] HttpRequestData request,
        CancellationToken cancellationToken)
    {
        var query = System.Web.HttpUtility.ParseQueryString(request.Url.Query);
        var dryRun = bool.TryParse(query["dryRun"], out var d) && d;
        var onlyProjectId = query["projectId"];

        // ?days=N widens the query window for a dry-run or targeted run (ignored on a normal real run
        // to avoid accidental reprocessing).
        DateTimeOffset? overrideSince = null;
        if ((dryRun || !string.IsNullOrWhiteSpace(onlyProjectId)) && int.TryParse(query["days"], out var days) && days > 0)
        {
            overrideSince = DateTimeOffset.UtcNow.AddDays(-days);
        }

        _logger.LogInformation("Manual ProjectSync run requested (dryRun={DryRun}, days={Days}, projectId={ProjectId}).",
            dryRun, query["days"] ?? "-", onlyProjectId ?? "-");

        try
        {
            var result = await _processor.RunAsync(
                new RunOptions { DryRun = dryRun, OverrideSince = overrideSince, OnlyProjectId = onlyProjectId }, cancellationToken);

            var response = request.CreateResponse(HttpStatusCode.OK);
            await response.WriteAsJsonAsync(result, cancellationToken);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Manual ProjectSync run failed.");
            var response = request.CreateResponse(HttpStatusCode.InternalServerError);
            await response.WriteAsJsonAsync(new { error = ex.Message }, cancellationToken);
            return response;
        }
    }
}
