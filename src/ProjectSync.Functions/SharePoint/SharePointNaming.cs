namespace ProjectSync.SharePoint;

/// <summary>Helpers for producing SharePoint-safe folder/document-set leaf names.</summary>
public static class SharePointNaming
{
    // Characters SharePoint disallows in file/folder leaf names.
    private static readonly char[] InvalidNameChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '#', '%' };

    /// <summary>
    /// Builds a document-set folder name as "{first N chars of customer name} | {project id}",
    /// sanitized for SharePoint (the '|' is illegal and becomes '-'). Falls back to just the project id
    /// when the customer name is blank. Because the (unique) project id is part of the name, names are
    /// effectively unique.
    /// </summary>
    public static string BuildDocumentSetName(string? customerName, string projectId, int customerMaxLength)
    {
        var customer = (customerName ?? string.Empty).Trim();
        if (customerMaxLength > 0 && customer.Length > customerMaxLength)
        {
            customer = customer[..customerMaxLength].Trim();
        }

        var raw = string.IsNullOrEmpty(customer) ? projectId : $"{customer} | {projectId}";
        return SanitizeLeafName(raw);
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
