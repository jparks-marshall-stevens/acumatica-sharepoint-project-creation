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

    /// <summary>
    /// Returns the distinct team-member emails for a project, from the team GI. Empty if no team GI
    /// is configured or the project has no team.
    /// </summary>
    Task<IReadOnlyList<string>> GetTeamEmailsAsync(string projectId, CancellationToken cancellationToken);

    /// <summary>
    /// Returns every team row (project id + email + modified date) from the team GI. The reconcile
    /// derives both "which projects changed since the watermark" and each project's full current team
    /// from this single read. Empty if no team GI is configured.
    /// </summary>
    Task<IReadOnlyList<TeamMemberRow>> GetTeamRowsAsync(CancellationToken cancellationToken);

    /// <summary>
    /// Writes the dataroom + client-upload URLs to the project's DataUrl/ClientUrl attributes via the
    /// contract-based REST API. No-op when the attribute IDs aren't configured. Returns true on success.
    /// </summary>
    Task<bool> WriteProjectUrlsAsync(string projectId, string? dataUrl, string? clientUrl, CancellationToken cancellationToken);
}

/// <summary>One row from the team GI: a person on a project, with the link's last-modified date.</summary>
public sealed record TeamMemberRow(string ProjectId, string Email, DateTimeOffset? ModifiedOn);
