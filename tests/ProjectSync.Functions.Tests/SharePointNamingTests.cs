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
    public void BuildDocumentSetName_FirstNCustomerChars_PipeProjectId()
    {
        // First 10 chars of customer + " | " + project id; the illegal '|' becomes '-'.
        var name = SharePointNaming.BuildDocumentSetName("Robert Palumbo", "10-31-21-74663", 10);
        Assert.Equal("Robert Pal - 10-31-21-74663", name);
    }

    [Fact]
    public void BuildDocumentSetName_ShortCustomer_UsedWhole()
    {
        var name = SharePointNaming.BuildDocumentSetName("GPM, Inc.", "15-31-26-10451", 10);
        Assert.Equal("GPM, Inc. - 15-31-26-10451", name);
    }

    [Fact]
    public void BuildDocumentSetName_TruncationTrimsTrailingSpace()
    {
        // "Kelleher &" is 10 chars; no trailing space here, but verify ampersand is kept and trim works.
        var name = SharePointNaming.BuildDocumentSetName("Kelleher & Holland, LLC", "10-31-21-74664", 10);
        Assert.Equal("Kelleher & - 10-31-21-74664", name);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void BuildDocumentSetName_BlankCustomer_FallsBackToProjectId(string? customer)
    {
        Assert.Equal("10-31-21-74655", SharePointNaming.BuildDocumentSetName(customer, "10-31-21-74655", 10));
    }
}
