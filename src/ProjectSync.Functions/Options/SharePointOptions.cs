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
    /// Number of leading characters of the customer name used in the document-set folder name,
    /// which is "{customer[..N]} | {project id}".
    /// </summary>
    public int DocumentSetNameMaxLength { get; set; } = 10;

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
    /// Internal name of a text column the reconcile uses to stash a hash of the last-applied state
    /// (metadata + grantee emails). The reconcile bulk-reads this column and only re-applies a set when
    /// its desired hash differs — so unchanged projects cost zero writes. Auto-created if missing.
    /// </summary>
    public string SignatureColumn { get; set; } = "ProjectSyncSig";

    /// <summary>
    /// True when <see cref="ProjectManagerColumn"/> is a Person/Group (People) field rather than text.
    /// When true, the value is resolved to a site user via EnsureUser (preferring the GI's PM email —
    /// see Acumatica:ProjectManagerEmailField) and set as a FieldUserValue. If the user can't be
    /// resolved, the field is left blank (the rest of the metadata still gets written).
    /// </summary>
    public bool ProjectManagerIsPersonColumn { get; set; }

    public string PracticeColumn { get; set; } = "Practice";

    // --- Scoping (HubSpot-sourced) workspace columns ---

    /// <summary>Internal name of the column holding the HubSpot deal id (scoping idempotency key). Auto-created.</summary>
    public string HubSpotDealIdColumn { get; set; } = "HubSpotDealId";

    /// <summary>Internal name of the column holding the workspace lifecycle status (Scoping/Active). Auto-created.</summary>
    public string StatusColumn { get; set; } = "Status";

    /// <summary>Status value stamped on a scoping-phase (HubSpot) workspace.</summary>
    public string ScopingStatusValue { get; set; } = "Scoping";

    /// <summary>Status value stamped on an Acumatica project workspace (the post-scoping / ERP phase).</summary>
    public string ProjectStatusValue { get; set; } = "Execution";

    // --- Client Uploads folder + external "Request files" (upload-only) sharing link ---

    /// <summary>
    /// When true, each newly-created document set gets a <see cref="ClientUploadsFolderName"/> subfolder
    /// and an anonymous upload-only ("Request files") sharing link is generated for it, stored in
    /// <see cref="ClientUploadLinkColumn"/>. Off by default: the Graph call requires the app to have a
    /// Microsoft Graph <c>Sites.Selected</c> grant on the site (separate from the SharePoint/CSOM grant),
    /// and the tenant must allow "Anyone" links for the anonymous scope. Fail-soft — if link creation
    /// fails, the document set is still created and the run succeeds. Lifecycle is create-once: the link
    /// is minted only when the set is first created, never refreshed by reconcile.
    /// </summary>
    public bool CreateClientUploadLink { get; set; }

    /// <summary>Name of the upload subfolder created inside each document set.</summary>
    public string ClientUploadsFolderName { get; set; } = "Client Uploads";

    /// <summary>
    /// Internal name of a text/URL column on the document set that receives the generated upload link.
    /// If the column is absent the link is still created (and logged) but not stamped.
    /// </summary>
    public string ClientUploadLinkColumn { get; set; } = "ClientUploadLink";

    /// <summary>Days until the upload link expires (Graph <c>expirationDateTime</c>). Default 30.</summary>
    public int ClientUploadLinkExpirationDays { get; set; } = 30;

    /// <summary>
    /// Sharing scope for the upload link: "anonymous" (Anyone with the link — required for external
    /// clients who have no Marshall &amp; Stevens login) or "organization" (only signed-in M&amp;S users).
    /// </summary>
    public string ClientUploadLinkScope { get; set; } = "anonymous";

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
