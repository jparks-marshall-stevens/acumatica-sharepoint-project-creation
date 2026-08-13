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

    /// <summary>
    /// Maximum length of the document-set folder name, taken from the start of the project description.
    /// Colliding names are made unique by appending the project id.
    /// </summary>
    public int DocumentSetNameMaxLength { get; set; } = 40;

    /// <summary>When true, set explicit permissions on each document set (project manager + practice leader).</summary>
    public bool SetProjectPermissions { get; set; } = true;

    /// <summary>SharePoint role/permission level granted to the PM and practice leader, e.g. "Edit" or "Contribute".</summary>
    public string PermissionLevel { get; set; } = "Edit";

    /// <summary>
    /// When true, break role inheritance so only the PM, the practice leader, and the site Owners can
    /// access the set. When false, those users are granted on top of the library's existing access.
    /// Requires the app to have Full Control on the site (Sites.Selected fullcontrol).
    /// </summary>
    public bool RestrictPermissions { get; set; } = true;

    // --- Internal (list) field names for the metadata columns on the document set ---
    // These are the *internal* names of the site/list columns. Adjust to match your library.
    public string ProjectIdColumn { get; set; } = "ProjectId";
    public string CustomerNameColumn { get; set; } = "CustomerName";
    public string ProjectNameColumn { get; set; } = "ProjectName";
    public string ProjectManagerColumn { get; set; } = "ProjectManager";

    /// <summary>
    /// True when <see cref="ProjectManagerColumn"/> is a Person/Group (People) field rather than text.
    /// When true, the value is resolved to a site user via EnsureUser (preferring the GI's PM email —
    /// see Acumatica:ProjectManagerEmailField) and set as a FieldUserValue. If the user can't be
    /// resolved, the field is left blank (the rest of the metadata still gets written).
    /// </summary>
    public bool ProjectManagerIsPersonColumn { get; set; }

    public string PracticeColumn { get; set; } = "Practice";

    /// <summary>
    /// Maps Acumatica practice values to SharePoint destinations. An indexed list (rather than a
    /// dictionary keyed by practice) so it binds cleanly from environment variables — practice
    /// values contain spaces and '&', which are invalid in config key names. Use Practice "*" as
    /// the catch-all/default.
    /// </summary>
    public List<PracticeMappingEntry> PracticeMappings { get; set; } = new();
}

/// <summary>Maps one practice value to where its document sets should be created.</summary>
public sealed class PracticeMappingEntry
{
    /// <summary>The Acumatica practice value this entry maps, or "*" for the catch-all default.</summary>
    public string Practice { get; set; } = string.Empty;

    /// <summary>
    /// Email/UPN of the practice leader granted access to every document set for this practice
    /// (resolved via EnsureUser). Optional.
    /// </summary>
    public string? PracticeLeaderEmail { get; set; }

    /// <summary>Optional site override. If null/empty, <see cref="SharePointOptions.SiteUrl"/> is used.</summary>
    public string? SiteUrl { get; set; }

    /// <summary>Document library (list) title, e.g. "Documents".</summary>
    public string Library { get; set; } = string.Empty;

    /// <summary>
    /// Optional parent folder path within the library (e.g. "Projects/Active"). Folders are ensured
    /// (created if missing). Leave empty to place the set at the library root.
    /// </summary>
    public string? ParentFolder { get; set; }
}
