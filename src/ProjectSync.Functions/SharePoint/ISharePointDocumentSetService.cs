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
}

public sealed record DocumentSetResult(bool Created, string ServerRelativeUrl);

/// <summary>The intended destination for a project's Document Set (config-resolved, not yet created).</summary>
public sealed record DocumentSetPlan(string SiteUrl, string Library, string? ParentFolder, string SetName);
