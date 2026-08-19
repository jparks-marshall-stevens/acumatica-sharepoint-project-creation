using Microsoft.Extensions.Logging;
using ProjectSync.SharePoint;
using ProjectSync.State;

namespace ProjectSync;

/// <summary>
/// Orchestrates the client-upload notification cycle: reads a persisted watermark, scans for files added
/// to Client Uploads folders since then, and advances the watermark. Moving-forward only — the first run
/// just establishes the baseline (no backfill of historical uploads).
/// </summary>
public sealed class ClientUploadProcessor
{
    /// <summary>Named watermark for the client-upload scan (stored alongside the other cursors).</summary>
    public const string WatermarkName = "client-uploads";

    private readonly ISharePointDocumentSetService _sharePoint;
    private readonly ILastRunStore _store;
    private readonly TimeProvider _timeProvider;
    private readonly ILogger<ClientUploadProcessor> _logger;

    public ClientUploadProcessor(
        ISharePointDocumentSetService sharePoint,
        ILastRunStore store,
        TimeProvider timeProvider,
        ILogger<ClientUploadProcessor> logger)
    {
        _sharePoint = sharePoint;
        _store = store;
        _timeProvider = timeProvider;
        _logger = logger;
    }

    public async Task<ClientUploadScanResult> RunAsync(CancellationToken cancellationToken)
    {
        // Capture the cut-off BEFORE scanning, so files uploaded during the scan are caught next cycle
        // (rather than double-reported).
        var now = _timeProvider.GetUtcNow();
        var watermark = await _store.GetWatermarkAsync(WatermarkName, cancellationToken);

        if (watermark is null)
        {
            await _store.SetWatermarkAsync(WatermarkName, now, cancellationToken);
            _logger.LogInformation("Client-upload scan first run: baseline set to {Now:o}; no historical backfill.", now);
            return new ClientUploadScanResult();
        }

        var result = await _sharePoint.ScanAndNotifyClientUploadsAsync(watermark.Value, cancellationToken);
        await _store.SetWatermarkAsync(WatermarkName, now, cancellationToken);
        return result;
    }
}
