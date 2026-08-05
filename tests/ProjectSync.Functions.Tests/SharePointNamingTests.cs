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
}
