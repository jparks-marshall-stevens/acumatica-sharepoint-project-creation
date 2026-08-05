using ProjectSync.Acumatica;

namespace ProjectSync.SharePoint;

public interface ISharePointDocumentSetService
{
    /// <summary>
    /// Ensures a Document Set exists for the given project (creating it if absent) and stamps
    /// the Project Id / Customer Name / Project Name / Project Manager metadata. Idempotent.
    /// </summary>
    Task<DocumentSetResult> EnsureProjectDocumentSetAsync(AcumaticaProject project, CancellationToken cancellationToken);
}

public sealed record DocumentSetResult(bool Created, string ServerRelativeUrl);
