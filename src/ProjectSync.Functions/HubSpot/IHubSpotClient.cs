namespace ProjectSync.HubSpot;

public interface IHubSpotClient
{
    /// <summary>
    /// Returns deals whose last-modified timestamp is strictly greater than
    /// <paramref name="modifiedAfterUtc"/>, filtered to the configured pipeline/stages, ordered
    /// oldest-modified first. Pages up to <paramref name="maxResults"/> and never past HubSpot's
    /// 10,000-result search window (the remainder is picked up on the next poll as the watermark advances).
    /// </summary>
    Task<IReadOnlyList<HubSpotDeal>> GetDealsModifiedAfterAsync(
        DateTimeOffset modifiedAfterUtc,
        int maxResults,
        CancellationToken cancellationToken);

    /// <summary>
    /// Resolves the customer name for a deal from its client contact: the contact's "company" text →
    /// its associated company name → the deal name (fallback). Returns the deal name if no client
    /// contact is present.
    /// </summary>
    Task<string?> ResolveCustomerNameAsync(HubSpotDeal deal, CancellationToken cancellationToken);

    /// <summary>Returns a map of HubSpot owner id → email, for resolving deal owners.</summary>
    Task<IReadOnlyDictionary<string, string>> GetOwnerEmailsAsync(CancellationToken cancellationToken);
}
