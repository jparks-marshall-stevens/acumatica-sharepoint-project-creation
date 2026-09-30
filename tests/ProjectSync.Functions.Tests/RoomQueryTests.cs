using ProjectSync.SharePoint;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class RoomQueryTests
{
    [Fact]
    public void Lookup_LeadsWithTheKeyColumn_SoItsIndexCanBeUsed()
    {
        var caml = RoomQuery.Lookup("Project_x0020_Id", "10-11-11-26290", 1, new[] { "FileRef" });
        var key = caml.IndexOf("Project_x0020_Id", StringComparison.Ordinal);
        var folder = caml.IndexOf("FSObjType", StringComparison.Ordinal);
        Assert.True(key >= 0 && key < folder, caml);
        Assert.Contains("<RowLimit>1</RowLimit>", caml);
    }

    [Fact]
    public void Lookup_EscapesTheValue()
    {
        var caml = RoomQuery.Lookup("OpportunityId", "A&B<1>", 2, Array.Empty<string>());
        Assert.Contains("A&amp;B&lt;1&gt;", caml);
    }

    [Fact]
    public void FolderChildren_IsUnfilteredIdOrderedAndPaged_NotRecursive()
    {
        var caml = RoomQuery.FolderChildren(new[] { "Project_x0020_Id" });
        Assert.DoesNotContain("<Where>", caml);
        Assert.DoesNotContain("RecursiveAll", caml);
        Assert.Contains("<OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>", caml);
        Assert.Contains($"<RowLimit Paged='TRUE'>{RoomQuery.PageSize}</RowLimit>", caml);
        Assert.Contains("<FieldRef Name='ContentTypeId'/>", caml);
        Assert.True(RoomQuery.PageSize < 5000);
    }

    [Theory]
    [InlineData("1", "0x0120D52000A1B2", true)]
    [InlineData("1", "0x0120", false)]                // plain folder (e.g. MN ESOP Clients)
    [InlineData("0", "0x0120D52000A1B2", false)]      // a file
    [InlineData("1", null, false)]
    public void IsRoom_OnlyMatchesDocumentSetFolders(string fsObjType, string? contentTypeId, bool expected) =>
        Assert.Equal(expected, RoomQuery.IsRoom(fsObjType, contentTypeId));

    [Theory]
    [InlineData("/sites/ESOP/Shared Documents", "Projects/Current", "/sites/ESOP/Shared Documents/Projects/Current")]
    [InlineData("/sites/ESOP/Shared Documents/", "/Projects/Current/", "/sites/ESOP/Shared Documents/Projects/Current")]
    [InlineData("/sites/ESOP/Shared Documents", null, "/sites/ESOP/Shared Documents")]
    public void ParentFolderUrl_JoinsRootAndParent(string root, string? parent, string expected) =>
        Assert.Equal(expected, RoomQuery.ParentFolderUrl(root, parent));
}
