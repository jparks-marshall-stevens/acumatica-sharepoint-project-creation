namespace ProjectSync.SharePoint;

/// <summary>Helpers for producing SharePoint-safe folder/document-set leaf names.</summary>
public static class SharePointNaming
{
    // Characters SharePoint disallows in file/folder leaf names.
    private static readonly char[] InvalidNameChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '#', '%' };

    /// <summary>
    /// Builds a document-set folder name as "{first N chars of customer name} ({project id})",
    /// sanitized for SharePoint. Falls back to just the project id when the customer name is blank.
    /// Because the (unique) project id is part of the name, names are effectively unique.
    /// </summary>
    public static string BuildDocumentSetName(string? customerName, string projectId, int customerMaxLength)
    {
        var customer = (customerName ?? string.Empty).Trim();
        if (customerMaxLength > 0 && customer.Length > customerMaxLength)
        {
            customer = customer[..customerMaxLength].Trim();
        }

        var raw = string.IsNullOrEmpty(customer) ? projectId : $"{customer} ({projectId})";
        return SanitizeLeafName(raw);
    }

    /// <summary>
    /// Builds an absolute URL to a document set / folder from the site URL and either a server-relative
    /// path or an already-absolute URL. When the input is absolute (as <c>DocumentSet.Create</c> can
    /// return), only its path is kept — otherwise the site origin would be prepended to a full URL and the
    /// host would be doubled (<c>…sharepoint.com…sharepoint.com/sites/…</c>). Spaces are percent-encoded.
    /// </summary>
    public static string ToAbsoluteUrl(string siteUrl, string urlOrServerRelative)
    {
        var origin = new Uri(siteUrl).GetLeftPart(UriPartial.Authority);
        var path = urlOrServerRelative.StartsWith("http", StringComparison.OrdinalIgnoreCase)
            ? Uri.UnescapeDataString(new Uri(urlOrServerRelative).AbsolutePath)
            : urlOrServerRelative;
        return origin + path.Replace(" ", "%20");
    }

    /// <summary>
    /// Replaces characters SharePoint forbids in leaf names with '-', trims surrounding
    /// whitespace and dots, and falls back to "Untitled" for an otherwise-empty result.
    /// </summary>
    public static string SanitizeLeafName(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return "Untitled";
        }

        var cleaned = new string(name.Select(c => InvalidNameChars.Contains(c) ? '-' : c).ToArray());
        cleaned = cleaned.Trim().Trim('.').Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "Untitled" : cleaned;
    }
}
