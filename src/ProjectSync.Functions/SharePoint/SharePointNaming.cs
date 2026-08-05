namespace ProjectSync.SharePoint;

/// <summary>Helpers for producing SharePoint-safe folder/document-set leaf names.</summary>
public static class SharePointNaming
{
    // Characters SharePoint disallows in file/folder leaf names.
    private static readonly char[] InvalidNameChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '#', '%' };

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
