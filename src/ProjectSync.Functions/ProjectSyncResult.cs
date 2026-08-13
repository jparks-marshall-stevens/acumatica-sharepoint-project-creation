namespace ProjectSync;

/// <summary>Outcome of a single sync cycle.</summary>
public sealed record ProjectSyncResult
{
    /// <summary>True when this was a preview run: nothing was created and the watermark was not moved.</summary>
    public bool DryRun { get; init; }

    /// <summary>Number of projects returned by the Generic Inquiry.</summary>
    public int Found { get; init; }

    /// <summary>Document sets newly created this cycle (0 in a dry run).</summary>
    public int Created { get; init; }

    /// <summary>Existing document sets whose metadata was refreshed (0 in a dry run).</summary>
    public int Updated { get; init; }

    /// <summary>In a dry run, the number of document sets that would be created/updated.</summary>
    public int Planned { get; init; }

    /// <summary>Projects skipped (excluded id or practice not in the allow-list).</summary>
    public int Skipped { get; init; }

    /// <summary>Set true when the cycle halted early because a project failed.</summary>
    public bool HadFailure { get; init; }

    /// <summary>The watermark at the end of the cycle (persisted only on a non-dry run).</summary>
    public DateTimeOffset Watermark { get; init; }

    /// <summary>Dry-run only: the document sets that would be created and where.</summary>
    public IReadOnlyList<PlannedDocumentSet>? Plan { get; init; }
}

/// <summary>A single projected document-set creation, surfaced by dry-run.</summary>
public sealed record PlannedDocumentSet
{
    public required string ProjectId { get; init; }
    public string? ProjectName { get; init; }
    public string? CustomerName { get; init; }
    public string? ProjectManager { get; init; }
    public string? ProjectManagerEmail { get; init; }
    public string? Practice { get; init; }
    public DateTimeOffset? CreatedDateTime { get; init; }
    public string? TargetSiteUrl { get; init; }
    public string? TargetLibrary { get; init; }
    public string? TargetFolder { get; init; }
    public string? DocumentSetName { get; init; }
}
