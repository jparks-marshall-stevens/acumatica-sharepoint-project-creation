namespace ProjectSync.SharePoint;

/// <summary>
/// Builds the CAML for finding rooms (document sets) so the queries keep working at any library size.
///
/// SharePoint refuses a query that has to scan more than 5,000 items (the list view threshold) unless its
/// leading condition is on an indexed column. Practice libraries pass that quickly once client files move
/// in (ESOP reached 25,000+ items with the MN ESOP migration), so the old library-wide
/// "FSObjType = folder AND …" queries stopped working. Two shapes stay safe:
/// <list type="bullet">
/// <item><see cref="Lookup"/>: one room by a key column. The key comes first; the column is indexed by the
/// sync, so SharePoint jumps straight to the matching rows however large the library is.</item>
/// <item><see cref="FolderChildren"/>: every direct child of the rooms' parent folder (Projects/Current),
/// with no filter, in ID order, paged. An unfiltered, ID-ordered, paged read is allowed at any size, and
/// scoping it to one folder means only the rooms themselves are read, not the files inside them.</item>
/// </list>
/// </summary>
public static class RoomQuery
{
    /// <summary>Content-type id prefix shared by every Document Set content type.</summary>
    public const string DocumentSetContentTypePrefix = "0x0120D520";

    /// <summary>Page size for <see cref="FolderChildren"/>; well under the 5,000 threshold.</summary>
    public const int PageSize = 2000;

    /// <summary>
    /// Folders anywhere in the library whose <paramref name="column"/> equals <paramref name="value"/>.
    /// The key condition leads so the column's index can satisfy it.
    /// </summary>
    public static string Lookup(string column, string value, int rowLimit, IEnumerable<string> viewFields)
    {
        var safeValue = System.Security.SecurityElement.Escape(value) ?? value;
        return
            "<View Scope='RecursiveAll'><Query><Where><And>" +
            $"<Eq><FieldRef Name='{column}'/><Value Type='Text'>{safeValue}</Value></Eq>" +
            "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
            "</And></Where></Query>" +
            ViewFields(viewFields) +
            $"<RowLimit>{rowLimit}</RowLimit></View>";
    }

    /// <summary>
    /// Every direct child of the folder the query is scoped to (set CamlQuery.FolderServerRelativeUrl),
    /// unfiltered and ID-ordered so it pages at any size. Filter to rooms with <see cref="IsRoom"/>.
    /// </summary>
    public static string FolderChildren(IEnumerable<string> viewFields) =>
        "<View><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query>" +
        ViewFields(viewFields.Concat(new[] { "FSObjType", "ContentTypeId" })) +
        $"<RowLimit Paged='TRUE'>{PageSize}</RowLimit></View>";

    /// <summary>True for a folder whose content type is a Document Set (a room, not a plain folder or file).</summary>
    public static bool IsRoom(object? fsObjType, object? contentTypeId) =>
        string.Equals(fsObjType?.ToString(), "1", StringComparison.Ordinal)
        && (contentTypeId?.ToString() ?? string.Empty).StartsWith(DocumentSetContentTypePrefix, StringComparison.OrdinalIgnoreCase);

    /// <summary>Server-relative URL of the rooms' parent folder: the library root plus the mapping's ParentFolder.</summary>
    public static string ParentFolderUrl(string libraryRootServerRelativeUrl, string? parentFolder)
    {
        var root = libraryRootServerRelativeUrl.TrimEnd('/');
        var rel = string.Join('/', (parentFolder ?? string.Empty)
            .Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(SharePointNaming.SanitizeLeafName));
        return rel.Length == 0 ? root : $"{root}/{rel}";
    }

    private static string ViewFields(IEnumerable<string> fields) =>
        "<ViewFields>" +
        string.Concat(fields.Where(f => !string.IsNullOrWhiteSpace(f)).Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(f => $"<FieldRef Name='{f}'/>")) +
        "</ViewFields>";
}
