namespace ProjectSync.Options;

/// <summary>Configuration for the last-run state store (Azure Blob).</summary>
public sealed class StateOptions
{
    public const string SectionName = "State";

    /// <summary>
    /// Blob storage connection string. Defaults to using the Functions "AzureWebJobsStorage"
    /// account when left empty (resolved in Program.cs).
    /// </summary>
    public string? ConnectionString { get; set; }

    public string ContainerName { get; set; } = "projectsync-state";

    public string BlobName { get; set; } = "last-run.txt";

    /// <summary>
    /// On the very first run (no stored state), how far back to look. Prevents importing the entire
    /// project history the first time the function executes.
    /// </summary>
    public int FirstRunLookbackHours { get; set; } = 24;

    /// <summary>
    /// Safety overlap subtracted from the last-run watermark to avoid missing records created
    /// during the previous poll window (clock skew / commit latency). Dedup handles reprocessing.
    /// </summary>
    public int OverlapMinutes { get; set; } = 5;
}
