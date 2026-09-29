namespace ProjectSync.SharePoint;

/// <summary>Builds the CAML for the client-upload scan.</summary>
public static class ClientUploadQuery
{
    /// <summary>
    /// Files created after <paramref name="since"/>, library-wide. The Created filter comes FIRST: on a
    /// library over the 5,000-item list view threshold, SharePoint only allows the query when its leading
    /// condition is on an indexed column (Created is indexed by the scan) and narrows the rows below the limit.
    /// </summary>
    public static string Build(DateTimeOffset since)
    {
        var sinceUtc = since.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
        return
            "<View Scope='RecursiveAll'><Query><Where><And>" +
            $"<Gt><FieldRef Name='Created'/><Value Type='DateTime' IncludeTimeValue='TRUE'>{sinceUtc}</Value></Gt>" +
            "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq>" +
            "</And></Where></Query>" +
            "<ViewFields><FieldRef Name='FileRef'/><FieldRef Name='FileLeafRef'/></ViewFields>" +
            "<RowLimit Paged='TRUE'>1000</RowLimit></View>";
    }

    /// <summary>True when SharePoint refused a query for exceeding the 5,000-item list view threshold.</summary>
    public static bool IsListViewThreshold(string? serverErrorTypeName, string? message) =>
        (serverErrorTypeName?.Contains("SPQueryThrottledException", StringComparison.OrdinalIgnoreCase) ?? false)
        || (message?.Contains("list view threshold", StringComparison.OrdinalIgnoreCase) ?? false);

    /// <summary>
    /// Keeps files under a Client Uploads folder (<paramref name="marker"/>, e.g. "/Client Uploads/") and
    /// groups their leaf names by the owning document set — the path before the marker. Files in subfolders
    /// of Client Uploads belong to the same document set.
    /// </summary>
    public static Dictionary<string, List<string>> GroupByDocSet(
        IEnumerable<(string FileRef, string Leaf)> files, string marker)
    {
        var byDocSet = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (fileRef, leaf) in files)
        {
            var idx = fileRef.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                continue; // not under a Client Uploads folder
            }

            var docSetUrl = fileRef[..idx];
            if (!byDocSet.TryGetValue(docSetUrl, out var names))
            {
                names = new List<string>();
                byDocSet[docSetUrl] = names;
            }

            names.Add(leaf);
        }

        return byDocSet;
    }
}
