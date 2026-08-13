using ProjectSync.SharePoint;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class SharePointNamingTests
{
    [Theory]
    [InlineData("PR000123", "PR000123")]
    [InlineData("AB/CD", "AB-CD")]
    [InlineData("a:b*c?d\"e<f>g|h#i%j", "a-b-c-d-e-f-g-h-i-j")]
    [InlineData("  spaced  ", "spaced")]
    [InlineData("trailing.", "trailing")]
    [InlineData(".leading", "leading")]
    public void SanitizeLeafName_ReplacesInvalidCharsAndTrims(string input, string expected)
    {
        Assert.Equal(expected, SharePointNaming.SanitizeLeafName(input));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("...")]
    public void SanitizeLeafName_EmptyOrAllTrimmed_FallsBackToUntitled(string input)
    {
        Assert.Equal("Untitled", SharePointNaming.SanitizeLeafName(input));
    }

    [Fact]
    public void SanitizeLeafName_AllInvalidChars_BecomesDashes()
    {
        // Invalid characters are replaced (not removed), so a slash-only name is not "empty".
        Assert.Equal("---", SharePointNaming.SanitizeLeafName("///"));
    }

    [Fact]
    public void BuildDocumentSetName_TruncatesToMaxLength()
    {
        var description = "Fair market value of a 1.0% ownership interest as of June 30, 2026";
        var name = SharePointNaming.BuildDocumentSetName(description, "10-31-21-74655", 40);
        // First 40 chars, with the SharePoint-illegal '%' replaced by '-'.
        Assert.Equal("Fair market value of a 1.0- ownership in", name);
        Assert.Equal(40, name.Length);
    }

    [Fact]
    public void BuildDocumentSetName_SanitizesIllegalCharsInDescription()
    {
        var name = SharePointNaming.BuildDocumentSetName("Mom Holdings & Koddi: Valuation/Report", "P1", 40);
        Assert.Equal("Mom Holdings & Koddi- Valuation-Report", name); // ':' and '/' -> '-', '&' kept
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void BuildDocumentSetName_BlankDescription_FallsBackToProjectId(string? description)
    {
        Assert.Equal("10-31-21-74655", SharePointNaming.BuildDocumentSetName(description, "10-31-21-74655", 40));
    }

    [Fact]
    public void BuildDocumentSetName_ShortDescription_KeptAsIs()
    {
        Assert.Equal("Odessa Separator Valuation",
            SharePointNaming.BuildDocumentSetName("Odessa Separator Valuation", "17-34-19-11617", 40));
    }
}
