using ProjectSync.Acumatica;

namespace ProjectSync.SharePoint;

public interface ISharePointDocumentSetService
{
    /// <summary>
    /// Ensures a Document Set exists for the given project (creating it if absent) and stamps
    /// the Project Id / Customer Name / Project Name / Project Manager metadata. Idempotent.
    ///
    /// Lookup order: the Project Id column, then — when the project carries a HubSpot deal id — the
    /// deal-id column, so an engagement that started as a scoping workspace is PROMOTED in place instead
    /// of getting a second folder.
    /// </summary>
    Task<DocumentSetResult> EnsureProjectDocumentSetAsync(AcumaticaProject project, CancellationToken cancellationToken);

    /// <summary>
    /// Resolves where a project's Document Set would be created (site / library / folder / name)
    /// from configuration only — no SharePoint connection. Used by dry-run.
    /// </summary>
    DocumentSetPlan PlanDocumentSet(AcumaticaProject project);

    /// <summary>
    /// Ensures a scoping-phase Document Set exists for a HubSpot deal (created if absent), keyed on the
    /// HubSpot deal id, stamped Customer / Project name / HubSpotDealId / Status=Scoping, and granting the
    /// deal owner + practice leader. Idempotent.
    /// </summary>
    Task<DocumentSetResult> EnsureScopingWorkspaceAsync(ScopingWorkspace workspace, CancellationToken cancellationToken);

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

/// <summary>
/// Outcome of ensuring one document set. <paramref name="Promoted"/> marks the special case where an
/// existing SCOPING workspace was converted in place into the project workspace (no new folder, no move)
/// — reported separately because it is the one path that rewrites another phase’s metadata and access.
/// </summary>
public sealed record DocumentSetResult(bool Created, string ServerRelativeUrl, bool Promoted = false);

/// <summary>A scoping-phase workspace to create/ensure from a HubSpot deal.</summary>
public sealed record ScopingWorkspace
{
    public required string DealId { get; init; }
    public string? CustomerName { get; init; }
    public string? ProjectName { get; init; }
    public string? Practice { get; init; }

    /// <summary>Deal owner email — granted Edit alongside the practice leader.</summary>
    public string? OwnerEmail { get; init; }

    /// <summary>
    /// Human-facing opportunity number, refreshed on every poll. Not the idempotency key (the deal id is) —
    /// it is the value an Acumatica project's PQCode is matched against at promotion.
    /// </summary>
    public string? OpportunityId { get; init; }
}

/// <summary>The intended destination for a project's Document Set (config-resolved, not yet created).</summary>
public sealed record DocumentSetPlan(string SiteUrl, string Library, string? ParentFolder, string SetName);
