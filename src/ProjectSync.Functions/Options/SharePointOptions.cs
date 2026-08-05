namespace ProjectSync.Options;

/// <summary>
/// Configuration for SharePoint access (Azure AD app-only, certificate auth) and the
/// mapping of Acumatica "practice" values to the target library/folder.
/// </summary>
public sealed class SharePointOptions
{
    public const string SectionName = "SharePoint";

    /// <summary>Azure AD (Entra) tenant id (GUID) or domain, e.g. contoso.onmicrosoft.com.</summary>
    public string AzureAdTenant { get; set; } = string.Empty;

    /// <summary>Application (client) id of the Azure AD app registration.</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// How to load the app-only certificate. Either a base64 PFX blob in <see cref="CertificateBase64"/>
    /// (with <see cref="CertificatePassword"/>), or a thumbprint resolved from the current user/machine store.
    /// </summary>
    public string? CertificateBase64 { get; set; }
    public string? CertificatePassword { get; set; }
    public string? CertificateThumbprint { get; set; }

    /// <summary>
    /// Default site collection URL, e.g. https://contoso.sharepoint.com/sites/Projects.
    /// A practice mapping may override the site per practice.
    /// </summary>
    public string SiteUrl { get; set; } = string.Empty;

    /// <summary>
    /// Content type name of the Document Set to create. Must exist and be enabled in the target library.
    /// Defaults to the OOTB "Document Set".
    /// </summary>
    public string DocumentSetContentType { get; set; } = "Document Set";

    // --- Internal (list) field names for the metadata columns on the document set ---
    // These are the *internal* names of the site/list columns. Adjust to match your library.
    public string ProjectIdColumn { get; set; } = "ProjectId";
    public string CustomerNameColumn { get; set; } = "CustomerName";
    public string ProjectNameColumn { get; set; } = "ProjectName";
    public string ProjectManagerColumn { get; set; } = "ProjectManager";
    public string PracticeColumn { get; set; } = "Practice";

    /// <summary>
    /// Maps an Acumatica practice value to the SharePoint destination. Key = practice value.
    /// Use "*" as a catch-all/default when a practice has no explicit mapping.
    /// </summary>
    public Dictionary<string, PracticeDestination> PracticeMappings { get; set; } = new();
}

/// <summary>Where a document set for a given practice should be created.</summary>
public sealed class PracticeDestination
{
    /// <summary>Optional site override. If null/empty, <see cref="SharePointOptions.SiteUrl"/> is used.</summary>
    public string? SiteUrl { get; set; }

    /// <summary>Document library (list) title, e.g. "Project Documents".</summary>
    public string Library { get; set; } = string.Empty;

    /// <summary>
    /// Optional server-relative-ish parent folder path within the library (e.g. "Advisory/2026").
    /// Folders are ensured (created if missing). Leave empty to place the set at the library root.
    /// </summary>
    public string? ParentFolder { get; set; }
}
