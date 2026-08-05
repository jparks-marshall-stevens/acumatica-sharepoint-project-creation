namespace ProjectSync;

/// <summary>Outcome of a single sync cycle.</summary>
public sealed record ProjectSyncResult
{
    /// <summary>Number of projects returned by the Generic Inquiry.</summary>
    public int Found { get; init; }

    /// <summary>Document sets newly created this cycle.</summary>
    public int Created { get; init; }

    /// <summary>Existing document sets whose metadata was refreshed.</summary>
    public int Updated { get; init; }

    /// <summary>Set true when the cycle halted early because a project failed.</summary>
    public bool HadFailure { get; init; }

    /// <summary>The watermark persisted at the end of the cycle.</summary>
    public DateTimeOffset Watermark { get; init; }
}
