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

    // --- OAuth2 (client credentials) ---
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;

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
    public string PracticeField { get; set; } = "Practice";

    /// <summary>Http timeout for Acumatica calls, seconds.</summary>
    public int TimeoutSeconds { get; set; } = 100;
}
