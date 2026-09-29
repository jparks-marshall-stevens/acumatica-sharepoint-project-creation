using ProjectSync.SharePoint;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class ClientUploadQueryTests
{
    [Fact]
    public void Created_filter_leads_the_query_so_large_libraries_stay_under_the_threshold()
    {
        var caml = ClientUploadQuery.Build(new DateTimeOffset(2026, 9, 29, 5, 30, 2, TimeSpan.Zero));

        var created = caml.IndexOf("<FieldRef Name='Created'/>", StringComparison.Ordinal);
        var fsObjType = caml.IndexOf("<FieldRef Name='FSObjType'/>", StringComparison.Ordinal);
        Assert.True(created >= 0 && fsObjType >= 0);
        Assert.True(created < fsObjType, "The indexed Created condition must come first in the <And>.");
    }

    [Fact]
    public void Since_is_written_in_utc()
    {
        var central = new DateTimeOffset(2026, 9, 29, 0, 30, 2, TimeSpan.FromHours(-5));

        var caml = ClientUploadQuery.Build(central);

        Assert.Contains(">2026-09-29T05:30:02Z<", caml);
    }

    [Theory]
    [InlineData("Microsoft.SharePoint.SPQueryThrottledException", "anything", true)]
    [InlineData(null, "The attempted operation is prohibited because it exceeds the list view threshold.", true)]
    [InlineData("Microsoft.SharePoint.SPException", "Access denied.", false)]
    [InlineData(null, null, false)]
    public void Recognises_the_list_view_threshold_error(string? typeName, string? message, bool expected)
    {
        Assert.Equal(expected, ClientUploadQuery.IsListViewThreshold(typeName, message));
    }

    [Fact]
    public void Groups_uploads_by_document_set_including_subfolders_and_ignores_other_files()
    {
        const string room = "/sites/ESOP/Shared Documents/Projects/Current/Acme (18-27-12-1)";
        var files = new[]
        {
            ($"{room}/Client Uploads/a.pdf", "a.pdf"),
            ($"{room}/Client Uploads/2025/b.xlsx", "b.xlsx"),
            ($"{room}/Admin/engagement letter.docx", "engagement letter.docx"),
            ("/sites/ESOP/Shared Documents/Projects/Current/MN ESOP Clients/Toll Company/x.pdf", "x.pdf"),
        };

        var grouped = ClientUploadQuery.GroupByDocSet(files, "/Client Uploads/");

        var only = Assert.Single(grouped);
        Assert.Equal(room, only.Key);
        Assert.Equal(new[] { "a.pdf", "b.xlsx" }, only.Value);
    }
}
