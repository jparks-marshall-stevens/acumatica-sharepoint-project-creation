namespace ProjectSync.Options;

/// <summary>
/// Configuration for reading HubSpot deals (the scoping-phase source). Auth is a HubSpot
/// private-app access token (a bearer token) — no OAuth exchange needed. All values come from
/// app settings; put the token in local.settings.json (git-ignored) for local dev.
/// </summary>
public sealed class HubSpotOptions
{
    public const string SectionName = "HubSpot";

    /// <summary>HubSpot API base URL. Almost always the default.</summary>
    public string BaseUrl { get; set; } = "https://api.hubapi.com";

    // --- OAuth 2.0 (refresh-token grant) — the supported, durable auth path ---

    /// <summary>OAuth app Client ID (from your HubSpot developer account).</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>OAuth app Client Secret. Secret — never commit.</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Long-lived OAuth refresh token captured by the one-time HubSpotOAuthSetup tool. Secret — never
    /// commit. The token provider exchanges this for short-lived access tokens at runtime.
    /// </summary>
    public string RefreshToken { get; set; } = string.Empty;

    /// <summary>Redirect URI registered on the OAuth app; only used during the one-time setup flow.</summary>
    public string RedirectUri { get; set; } = "http://localhost:5127/callback";

    /// <summary>OAuth scopes to request (must match the app's configured scopes).</summary>
    public List<string> Scopes { get; set; } = new()
    {
        "crm.objects.deals.read",
        "crm.schemas.deals.read",
        "crm.objects.companies.read",
        "crm.objects.owners.read",
    };

    /// <summary>
    /// Optional static bearer token (a private-app token) used only if no OAuth refresh token is
    /// configured. Lets you fall back to a private app without code changes. Secret — never commit.
    /// </summary>
    public string AccessToken { get; set; } = string.Empty;

    /// <summary>
    /// Pipeline ids in scope for scoping workspaces. Empty = all pipelines. Get ids from the
    /// connectivity test (it lists pipelines + their internal ids).
    /// </summary>
    public List<string> PipelineIds { get; set; } = new();

    /// <summary>
    /// Terminal deal-stage ids (Won / Lost / Closed). A deal is "in scoping" until it reaches one of
    /// these, so deals in these stages are excluded from the poll. Empty = no stage filter.
    /// </summary>
    public List<string> TerminalStageIds { get; set; } = new();

    // --- Deal property mapping (HubSpot internal property names) ---

    /// <summary>Deal property holding the project/engagement name. Default HubSpot property is "dealname".</summary>
    public string DealNameProperty { get; set; } = "dealname";

    /// <summary>
    /// Optional deal property holding the customer name. Leave blank if the customer comes from the
    /// associated Company (association lookup is a later step); when blank, CustomerName is null.
    /// </summary>
    public string CustomerProperty { get; set; } = string.Empty;

    /// <summary>
    /// Deal property holding the practice/service line (drives the SharePoint destination). For M&amp;S this
    /// is "practices" — a multi-select whose values match the Acumatica taxonomy (e.g. "Estate &amp; Gift").
    /// </summary>
    public string PracticeProperty { get; set; } = string.Empty;

    /// <summary>
    /// Practice values to include (case-insensitive, contains-match against the multi-select
    /// <see cref="PracticeProperty"/>). Empty = all practices. e.g. ["Estate &amp; Gift"] to scope
    /// scoping-workspace creation to the Estate &amp; Gift practice.
    /// </summary>
    public List<string> IncludedPractices { get; set; } = new();

    /// <summary>Deal property holding the owner id. Default "hubspot_owner_id" (resolved to an email).</summary>
    public string OwnerIdProperty { get; set; } = "hubspot_owner_id";

    /// <summary>
    /// Deal property that points directly at the client contact. Default "client_contact_id". When blank
    /// on a deal, the client contact is found via the labeled association <see cref="ClientContactLabel"/>.
    /// </summary>
    public string ClientContactIdProperty { get; set; } = "client_contact_id";

    /// <summary>Deal→contact association label identifying the client contact. Default "Client Contact".</summary>
    public string ClientContactLabel { get; set; } = "Client Contact";

    /// <summary>
    /// When resolving the customer from the client contact: true = prefer the contact's "company" text
    /// field (always populated but sometimes messy), then the associated company name; false = prefer
    /// the associated company name (cleaner but sometimes blank), then the text field. Deal name is the
    /// final fallback either way.
    /// </summary>
    public bool CustomerCompanyTextFirst { get; set; } = true;

    /// <summary>Deal property holding the created timestamp. Default "createdate".</summary>
    public string CreatedProperty { get; set; } = "createdate";

    /// <summary>Deal property holding the last-modified timestamp (the poll watermark field). Default "hs_lastmodifieddate".</summary>
    public string ModifiedProperty { get; set; } = "hs_lastmodifieddate";

    /// <summary>Extra deal properties to also request (for discovery / future mapping). Optional.</summary>
    public List<string> ExtraProperties { get; set; } = new();

    /// <summary>Http timeout for HubSpot calls, seconds.</summary>
    public int TimeoutSeconds { get; set; } = 100;

    /// <summary>
    /// On the very first poll (no watermark yet), look back this many hours. Default 0 = moving-forward
    /// only: the first run stamps the watermark at "now" and processes nothing historical, so we don't
    /// pull the whole deal history. Raise it for a bounded one-time backfill.
    /// </summary>
    public int FirstRunLookbackHours { get; set; } = 0;

    /// <summary>
    /// Optional created-date floor: only deals whose CreatedAt is strictly after this are eligible for a
    /// scoping workspace. Because HubSpot bumps hs_lastmodifieddate constantly, a modified-date watermark
    /// alone would gradually pick up pre-existing open deals; this floor limits creation to deals created
    /// from go-live onward. Null = no floor (all in-scope deals eligible).
    /// </summary>
    public DateTimeOffset? CreatedAfter { get; set; }

    /// <summary>
    /// Small overlap (minutes) subtracted from the watermark on each poll, guarding against clock skew /
    /// modification-commit latency so a deal changed right at the boundary isn't missed.
    /// </summary>
    public int OverlapMinutes { get; set; } = 5;

    /// <summary>Max deals to pull in one poll (also bounded by HubSpot's 10,000-result search window).</summary>
    public int MaxDealsPerPoll { get; set; } = 10000;
}
