namespace ProjectSync.Options;

/// <summary>
/// Configuration for connecting to Acumatica and reading the projects Generic Inquiry.
/// All values are supplied via app settings (see local.settings.json for local dev).
/// </summary>
public sealed class AcumaticaOptions
{
    public const string SectionName = "Acumatica";

    /// <summary>Base URL of the Acumatica instance, e.g. https://mycompany.acumatica.com </summary>
    public string BaseUrl { get; set; } = string.Empty;

    /// <summary>Tenant/company login name (the value you pick on the sign-in screen).</summary>
    public string Tenant { get; set; } = string.Empty;

    // --- OAuth2 ---
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// OAuth grant type. "password" = Resource Owner Password Credentials (ROPC), which requires
    /// <see cref="Username"/>/<see cref="Password"/> and a real Acumatica service account.
    /// "client_credentials" = machine-to-machine (only if the connected app supports it).
    /// </summary>
    public string GrantType { get; set; } = "password";

    /// <summary>Acumatica service-account user name (ROPC only).</summary>
    public string Username { get; set; } = string.Empty;

    /// <summary>Acumatica service-account password (ROPC only).</summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>OAuth scope. Acumatica typically uses "api" (optionally "offline_access").</summary>
    public string Scope { get; set; } = "api";

    /// <summary>
    /// Name of the Generic Inquiry exposed over OData, e.g. "P-ProjectsCreated".
    /// Read from {BaseUrl}/t/{Tenant}/api/odata/gi/{GenericInquiryName}.
    /// </summary>
    public string GenericInquiryName { get; set; } = string.Empty;

    /// <summary>
    /// Field name (as it appears in the GI OData output) holding the project's created date/time.
    /// Used to build the $filter for "created since last run". e.g. "CreatedDateTime".
    /// </summary>
    public string CreatedDateTimeField { get; set; } = "CreatedDateTime";

    // --- Field mapping: GI column name -> the value we care about ---
    public string ProjectIdField { get; set; } = "ProjectID";
    public string ProjectNameField { get; set; } = "ProjectName";
    public string CustomerNameField { get; set; } = "CustomerName";
    public string ProjectManagerField { get; set; } = "ProjectManager";

    /// <summary>
    /// Optional GI column holding the project manager's email/UPN. Required only when the SharePoint
    /// Project Manager column is a People field (see SharePoint:ProjectManagerIsPersonColumn), so the
    /// person can be resolved reliably. Leave blank if the PM column is plain text.
    /// </summary>
    public string ProjectManagerEmailField { get; set; } = string.Empty;

    public string PracticeField { get; set; } = "Practice";

    /// <summary>
    /// Optional GI column holding the HubSpot identifier this project was converted from — the instance
    /// uses "PQCode", which carries the HubSpot opportunity number. Conversion is manual, so this value is
    /// the only thread linking the two systems; it lets the sync PROMOTE an existing scoping workspace in
    /// place instead of creating a second one. The value is matched against the opportunity-number column
    /// first and the HubSpot deal-id column second, so either identifier links. Leave blank to disable the
    /// promotion path; a blank value on an individual row just means "no scoping folder".
    /// </summary>
    public string HubSpotLinkField { get; set; } = string.Empty;

    // --- Team GI (project employees, for permission sync) ---
    /// <summary>Name of the OData GI returning one row per (project, employee) with the employee email.</summary>
    public string TeamGenericInquiryName { get; set; } = string.Empty;
    public string TeamProjectIdField { get; set; } = "ProjectID";
    public string TeamEmailField { get; set; } = "EmployeeEmail";
    public string TeamModifiedField { get; set; } = "LastModifiedDateTime";

    /// <summary>Http timeout for Acumatica calls, seconds.</summary>
    public int TimeoutSeconds { get; set; } = 100;

    /// <summary>
    /// Optional allow-list of practice values to process (case-insensitive). When empty, all
    /// practices are processed. Projects whose practice is not listed are skipped but still
    /// advance the watermark, so they are not re-read every cycle.
    /// e.g. ["Estate &amp; Gift"] to only create document sets for Estate &amp; Gift projects.
    /// </summary>
    public List<string> IncludedPractices { get; set; } = new();

    /// <summary>
    /// Project IDs to always ignore (case-insensitive), regardless of practice. Use for internal
    /// / system project monikers such as "X" (Non-Project Code). Excluded projects are skipped but
    /// still advance the watermark.
    /// </summary>
    public List<string> ExcludedProjectIds { get; set; } = new();
}
