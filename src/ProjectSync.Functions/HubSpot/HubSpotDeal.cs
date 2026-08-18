namespace ProjectSync.HubSpot;

/// <summary>A single deal read from HubSpot — the scoping-phase source record.</summary>
public sealed record HubSpotDeal
{
    /// <summary>HubSpot deal object id — the stable idempotency key for the scoping workspace.</summary>
    public required string DealId { get; init; }

    public string? DealName { get; init; }
    public string? CustomerName { get; init; }
    public string? Practice { get; init; }

    public string? PipelineId { get; init; }
    public string? StageId { get; init; }

    /// <summary>HubSpot owner id (resolved to an email in a later step).</summary>
    public string? OwnerId { get; init; }
    public string? OwnerEmail { get; init; }

    /// <summary>Id of the deal's client contact (from the client-contact-id property), if set.</summary>
    public string? ClientContactId { get; init; }

    /// <summary>
    /// Human-facing opportunity number (HubSpot:OpportunityIdProperty) — the value a person types into the
    /// Acumatica PQCode field at conversion, and therefore the correlation key for promotion. Null when the
    /// property isn't configured or isn't set on this deal yet.
    /// </summary>
    public string? OpportunityId { get; init; }

    public DateTimeOffset? CreatedAt { get; init; }
    public DateTimeOffset? ModifiedAt { get; init; }

    /// <summary>All raw properties returned for the deal (name → value). Used for discovery/mapping.</summary>
    public IReadOnlyDictionary<string, string?> Properties { get; init; } =
        new Dictionary<string, string?>();
}
