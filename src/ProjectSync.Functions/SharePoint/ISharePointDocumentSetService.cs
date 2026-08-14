using ProjectSync.Acumatica;

namespace ProjectSync.SharePoint;

public interface ISharePointDocumentSetService
{
    /// <summary>
    /// Ensures a Document Set exists for the given project (creating it if absent) and stamps
    /// the Project Id / Customer Name / Project Name / Project Manager metadata. Idempotent.
    /// </summary>
    Task<DocumentSetResult> EnsureProjectDocumentSetAsync(AcumaticaProject project, CancellationToken cancellationToken);

    /// <summary>
    /// Resolves where a project's Document Set would be created (site / library / folder / name)
    /// from configuration only — no SharePoint connection. Used by dry-run.
    /// </summary>
    DocumentSetPlan PlanDocumentSet(AcumaticaProject project);

    /// <summary>
    /// Reconciles metadata + permissions of ALREADY-TRACKED document sets against the desired state.
    /// Only sets whose signature changed are re-applied (a single bulk read finds them). Projects
    /// without an existing set are skipped (no backfill). When <paramref name="onlyProjectIds"/> is
    /// non-null, only those ids are considered (incremental); null reconciles all tracked (daily sweep).
    /// </summary>
    Task<ReconcileResult> ReconcileAsync(
        IReadOnlyList<AcumaticaProject> desiredProjects,
        IReadOnlySet<string>? onlyProjectIds,
        CancellationToken cancellationToken);
}

/// <summary>Outcome of a reconcile pass.</summary>
public sealed record ReconcileResult
{
    public int Considered { get; init; }
    public int Updated { get; init; }
    public int Unchanged { get; init; }
    public int NotTracked { get; init; }
}

public sealed record DocumentSetResult(bool Created, string ServerRelativeUrl);

/// <summary>The intended destination for a project's Document Set (config-resolved, not yet created).</summary>
public sealed record DocumentSetPlan(string SiteUrl, string Library, string? ParentFolder, string SetName);
