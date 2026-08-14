namespace ProjectSync.State;

public interface ILastRunStore
{
    /// <summary>Returns the stored last-run (create) watermark, or null if this is the first run.</summary>
    Task<DateTimeOffset?> GetLastRunAsync(CancellationToken cancellationToken);

    /// <summary>Persists the create watermark (created date/time of the newest processed project).</summary>
    Task SetLastRunAsync(DateTimeOffset value, CancellationToken cancellationToken);

    /// <summary>Returns a named watermark (e.g. the reconcile team-modified cursor), or null if unset.</summary>
    Task<DateTimeOffset?> GetWatermarkAsync(string name, CancellationToken cancellationToken);

    /// <summary>Persists a named watermark.</summary>
    Task SetWatermarkAsync(string name, DateTimeOffset value, CancellationToken cancellationToken);
}
