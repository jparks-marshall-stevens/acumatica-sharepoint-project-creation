using System.Globalization;
using Azure;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.State;

/// <summary>
/// Stores the last-run watermark as an ISO-8601 UTC timestamp in a single Azure Blob.
/// </summary>
public sealed class BlobLastRunStore : ILastRunStore
{
    private readonly BlobClient _blob;
    private readonly ILogger<BlobLastRunStore> _logger;

    public BlobLastRunStore(BlobServiceClient serviceClient, IOptions<StateOptions> options, ILogger<BlobLastRunStore> logger)
    {
        var opts = options.Value;
        var container = serviceClient.GetBlobContainerClient(opts.ContainerName);
        container.CreateIfNotExists();
        _blob = container.GetBlobClient(opts.BlobName);
        _logger = logger;
    }

    public async Task<DateTimeOffset?> GetLastRunAsync(CancellationToken cancellationToken)
    {
        try
        {
            var result = await _blob.DownloadContentAsync(cancellationToken);
            var text = result.Value.Content.ToString().Trim();
            if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var value))
            {
                return value;
            }

            _logger.LogWarning("Last-run blob contained unparseable value '{Text}'; treating as first run.", text);
            return null;
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
    }

    public async Task SetLastRunAsync(DateTimeOffset value, CancellationToken cancellationToken)
    {
        var text = value.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
        await _blob.UploadAsync(BinaryData.FromString(text), overwrite: true, cancellationToken);
        _logger.LogDebug("Persisted last-run watermark {Value:o}", value);
    }
}
