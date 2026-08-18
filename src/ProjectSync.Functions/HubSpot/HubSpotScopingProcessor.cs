using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

namespace ProjectSync.HubSpot;

/// <summary>
/// Polls HubSpot for deals modified since a persisted watermark (so each run pulls only the delta, not
/// the whole deal history) and creates/updates a scoping SharePoint workspace for each in-scope deal,
/// keyed on the HubSpot deal id. The watermark advances to the newest processed modification time.
/// </summary>
public sealed class HubSpotScopingProcessor
{
    /// <summary>Named watermark for the HubSpot deal poll (stored alongside the other cursors).</summary>
    public const string DealWatermarkName = "hubspot-deals";

    private readonly IHubSpotClient _hubspot;
    private readonly ISharePointDocumentSetService _sharePoint;
    private readonly ILastRunStore _store;
    private readonly HubSpotOptions _options;
    private readonly TimeProvider _timeProvider;
    private readonly ILogger<HubSpotScopingProcessor> _logger;

    public HubSpotScopingProcessor(
        IHubSpotClient hubspot,
        ISharePointDocumentSetService sharePoint,
        ILastRunStore store,
        IOptions<HubSpotOptions> options,
        TimeProvider timeProvider,
        ILogger<HubSpotScopingProcessor> logger)
    {
        _hubspot = hubspot;
        _sharePoint = sharePoint;
        _store = store;
        _options = options.Value;
        _timeProvider = timeProvider;
        _logger = logger;
    }

    public async Task<HubSpotPollResult> RunAsync(bool dryRun, CancellationToken cancellationToken)
    {
        var watermark = await _store.GetWatermarkAsync(DealWatermarkName, cancellationToken);
        var nowUtc = _timeProvider.GetUtcNow();

        // First run: look back FirstRunLookbackHours (0 = moving-forward only). Later runs: apply a small
        // overlap so a deal modified right at the boundary isn't skipped.
        var queryFrom = watermark is null
            ? nowUtc.AddHours(-_options.FirstRunLookbackHours)
            : watermark.Value.AddMinutes(-_options.OverlapMinutes);

        _logger.LogInformation(
            "HubSpot poll starting{DryRun}. Watermark={Watermark:o}, querying modified > {From:o}",
            dryRun ? " (DRY RUN)" : string.Empty, watermark, queryFrom);

        var deals = await _hubspot.GetDealsModifiedAfterAsync(queryFrom, _options.MaxDealsPerPoll, cancellationToken);

        // Practice scope (client-side, contains-match against the multi-select practices property),
        // plus an optional created-date floor so pre-existing open deals aren't backfilled just because
        // HubSpot bumped their modified date.
        var included = _options.IncludedPractices;
        var floor = _options.CreatedAfter;
        var inScope = deals.Where(d =>
                (included.Count == 0 || included.Any(p => (d.Practice ?? string.Empty).Contains(p, StringComparison.OrdinalIgnoreCase))) &&
                (floor is null || (d.CreatedAt is { } created && created > floor)))
            .ToList();

        // Process oldest-modified first so a mid-batch failure holds the watermark just before the
        // offending deal (the rest retries next cycle).
        var ordered = inScope.OrderBy(d => d.ModifiedAt ?? queryFrom).ToList();
        var owners = await _hubspot.GetOwnerEmailsAsync(cancellationToken);
        var plan = new List<ScopingWorkspacePlan>();
        int created = 0, updated = 0;
        var hadFailure = false;
        DateTimeOffset? advanced = watermark; // how far the watermark can safely move on a partial failure

        foreach (var deal in ordered)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var customer = await _hubspot.ResolveCustomerNameAsync(deal, cancellationToken);
            var ownerEmail = deal.OwnerId is { } oid && owners.TryGetValue(oid, out var em) ? em : null;
            plan.Add(new ScopingWorkspacePlan
            {
                DealId = deal.DealId,
                CustomerName = customer,
                ProjectName = deal.DealName,
                Practice = deal.Practice,
                OwnerEmail = ownerEmail,
                StageId = deal.StageId,
            });

            if (!dryRun)
            {
                try
                {
                    var result = await _sharePoint.EnsureScopingWorkspaceAsync(new ScopingWorkspace
                    {
                        DealId = deal.DealId,
                        CustomerName = customer,
                        ProjectName = deal.DealName,
                        Practice = deal.Practice,
                        OwnerEmail = ownerEmail,
                        OpportunityId = deal.OpportunityId,
                    }, cancellationToken);

                    if (result.Created) created++; else updated++;
                }
                catch (Exception ex)
                {
                    hadFailure = true;
                    _logger.LogError(ex, "Failed to ensure scoping workspace for deal {DealId}; holding watermark for retry.", deal.DealId);
                    break;
                }
            }

            if (deal.ModifiedAt is { } m && (advanced is null || m > advanced))
            {
                advanced = m;
            }
        }

        // Advance the watermark — never regress. On a partial failure, hold at the last processed deal so
        // the remainder retries. On success, jump to the newest modification across ALL deals seen (so
        // out-of-scope deals aren't re-examined). When nothing changed, hold / establish the baseline.
        DateTimeOffset newWatermark;
        if (hadFailure)
        {
            var candidate = advanced ?? queryFrom;
            newWatermark = watermark is { } w && w > candidate ? w : candidate;
        }
        else if (deals.Count > 0)
        {
            var maxModified = deals.Max(d => d.ModifiedAt ?? queryFrom);
            newWatermark = watermark is { } w && w > maxModified ? w : maxModified;
        }
        else
        {
            newWatermark = watermark ?? nowUtc;
        }

        if (!dryRun)
        {
            await _store.SetWatermarkAsync(DealWatermarkName, newWatermark, cancellationToken);
        }

        _logger.LogInformation(
            "HubSpot poll complete{DryRun}. Modified={Found}, in-scope={InScope}, created={Created}, updated={Updated}, watermark {Old:o} → {New:o}.",
            dryRun ? " (DRY RUN)" : string.Empty, deals.Count, inScope.Count, created, updated, watermark, newWatermark);

        return new HubSpotPollResult
        {
            DryRun = dryRun,
            QueriedFrom = queryFrom,
            Found = deals.Count,
            InScope = inScope.Count,
            Created = created,
            Updated = updated,
            NewWatermark = newWatermark,
            Plan = plan,
        };
    }
}

/// <summary>Outcome of one HubSpot deal poll.</summary>
public sealed record HubSpotPollResult
{
    public bool DryRun { get; init; }
    public DateTimeOffset QueriedFrom { get; init; }
    public int Found { get; init; }
    public int InScope { get; init; }
    public int Created { get; init; }
    public int Updated { get; init; }
    public DateTimeOffset NewWatermark { get; init; }
    public IReadOnlyList<ScopingWorkspacePlan> Plan { get; init; } = Array.Empty<ScopingWorkspacePlan>();
}

/// <summary>What a scoping workspace would be created/updated with (dry-run before SharePoint writes).</summary>
public sealed record ScopingWorkspacePlan
{
    public required string DealId { get; init; }
    public string? CustomerName { get; init; }
    public string? ProjectName { get; init; }
    public string? Practice { get; init; }
    public string? OwnerEmail { get; init; }
    public string? StageId { get; init; }
}
