namespace ProjectSync.Acumatica;

public interface IAcumaticaClient
{
    /// <summary>
    /// Returns projects whose created date/time is strictly greater than <paramref name="createdAfterUtc"/>,
    /// read from the configured Generic Inquiry (OData feed), ordered oldest-first.
    /// </summary>
    Task<IReadOnlyList<AcumaticaProject>> GetProjectsCreatedAfterAsync(
        DateTimeOffset createdAfterUtc,
        CancellationToken cancellationToken);
}
