using System.Globalization;
using Azure;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.State;

/// <summary>
/// Stores watermarks as ISO-8601 UTC timestamps, one Azure Blob per watermark name.
/// </summary>
public sealed class BlobLastRunStore : ILastRunStore
{
    private readonly BlobContainerClient _container;
    private readonly string _lastRunBlobName;
    private readonly ILogger<BlobLastRunStore> _logger;

    public BlobLastRunStore(BlobServiceClient serviceClient, IOptions<StateOptions> options, ILogger<BlobLastRunStore> logger)
    {
        var opts = options.Value;
        _container = serviceClient.GetBlobContainerClient(opts.ContainerName);
        _container.CreateIfNotExists();
        _lastRunBlobName = opts.BlobName;
        _logger = logger;
    }

    public Task<DateTimeOffset?> GetLastRunAsync(CancellationToken cancellationToken)
        => GetAsync(_lastRunBlobName, cancellationToken);

    public Task SetLastRunAsync(DateTimeOffset value, CancellationToken cancellationToken)
        => SetAsync(_lastRunBlobName, value, cancellationToken);

    public Task<DateTimeOffset?> GetWatermarkAsync(string name, CancellationToken cancellationToken)
        => GetAsync(BlobNameFor(name), cancellationToken);

    public Task SetWatermarkAsync(string name, DateTimeOffset value, CancellationToken cancellationToken)
        => SetAsync(BlobNameFor(name), value, cancellationToken);

    private static string BlobNameFor(string name) => $"watermark-{name}.txt";

    private async Task<DateTimeOffset?> GetAsync(string blobName, CancellationToken cancellationToken)
    {
        try
        {
            var result = await _container.GetBlobClient(blobName).DownloadContentAsync(cancellationToken);
            var text = result.Value.Content.ToString().Trim();
            if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var value))
            {
                return value;
            }

            _logger.LogWarning("Watermark blob '{Blob}' contained unparseable value '{Text}'.", blobName, text);
            return null;
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
    }

    private async Task SetAsync(string blobName, DateTimeOffset value, CancellationToken cancellationToken)
    {
        var text = value.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
        await _container.GetBlobClient(blobName).UploadAsync(BinaryData.FromString(text), overwrite: true, cancellationToken);
        _logger.LogDebug("Persisted watermark '{Blob}' = {Value:o}", blobName, value);
    }
}
