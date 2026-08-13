namespace ProjectSync.SharePoint;

/// <summary>Helpers for producing SharePoint-safe folder/document-set leaf names.</summary>
public static class SharePointNaming
{
    // Characters SharePoint disallows in file/folder leaf names.
    private static readonly char[] InvalidNameChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '#', '%' };

    /// <summary>
    /// Builds a document-set folder name from the first <paramref name="maxLength"/> characters of the
    /// project description, sanitized for SharePoint. Falls back to the project id when the description
    /// is blank.
    /// </summary>
    public static string BuildDocumentSetName(string? description, string projectId, int maxLength)
    {
        var basis = string.IsNullOrWhiteSpace(description) ? projectId : description!;
        if (maxLength > 0 && basis.Length > maxLength)
        {
            basis = basis[..maxLength];
        }

        var name = SanitizeLeafName(basis);
        // If truncation/sanitization emptied it, fall back to the (sanitized) project id.
        return name == "Untitled" && !string.IsNullOrWhiteSpace(projectId)
            ? SanitizeLeafName(projectId)
            : name;
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
