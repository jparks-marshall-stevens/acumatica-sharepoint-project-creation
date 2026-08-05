namespace ProjectSync.State;

public interface ILastRunStore
{
    /// <summary>Returns the stored last-run watermark, or null if this is the first run.</summary>
    Task<DateTimeOffset?> GetLastRunAsync(CancellationToken cancellationToken);

    /// <summary>Persists the watermark (the created date/time of the newest processed project).</summary>
    Task SetLastRunAsync(DateTimeOffset value, CancellationToken cancellationToken);
}
