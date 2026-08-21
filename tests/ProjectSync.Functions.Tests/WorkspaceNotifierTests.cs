using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using ProjectSync.Notifications;
using ProjectSync.Options;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class WorkspaceNotifierTests
{
    private sealed class CapturingSender : IEmailSender
    {
        public List<EmailMessage> Sent { get; } = new();
        public Task SendAsync(EmailMessage message, CancellationToken cancellationToken)
        {
            Sent.Add(message);
            return Task.CompletedTask;
        }
    }

    private static WorkspaceNotice Notice() => new()
    {
        Phase = WorkspacePhase.Execution,
        CustomerName = "Acme Trust",
        EngagementName = "Gift valuation",
        IdLabel = "Project ID",
        IdValue = "10-31-21-70000",
        Practice = "Estate & Gift",
        DataroomUrl = "https://example.sharepoint.com/sites/GiftEstate/x",
        UploadLinkUrl = "https://example.sharepoint.com/:f:/s/GiftEstate/upload",
    };

    private static (WorkspaceNotifier N, CapturingSender S) Make(NotificationOptions? opts = null)
    {
        var sender = new CapturingSender();
        var notifier = new WorkspaceNotifier(
            sender,
            Microsoft.Extensions.Options.Options.Create(opts ?? new NotificationOptions { Enabled = true, BccAddress = "bcc@x.com" }),
            NullLogger<WorkspaceNotifier>.Instance);
        return (notifier, sender);
    }

    [Fact]
    public async Task Created_ExcludesLeader_Dedupes_AndSetsBcc()
    {
        var (n, s) = Make();
        await n.NotifyCreatedAsync(
            Notice(),
            new[] { "pm@x.com", "team@x.com", "PM@x.com", "leader@x.com" },
            leaderEmail: "leader@x.com",
            CancellationToken.None);

        var msg = Assert.Single(s.Sent);
        Assert.Equal(2, msg.To.Count);                    // pm + team, dedup case-insensitive
        Assert.DoesNotContain(msg.To, r => r.Equals("leader@x.com", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("bcc@x.com", msg.Bcc);
        Assert.Contains("New project dataroom", msg.Subject);
    }

    [Fact]
    public async Task Created_NoRecipientsAfterExclusion_SendsNothing()
    {
        var (n, s) = Make();
        await n.NotifyCreatedAsync(Notice(), new[] { "leader@x.com" }, "leader@x.com", CancellationToken.None);
        Assert.Empty(s.Sent);
    }

    [Fact]
    public async Task AccessAdded_UsesAddedSubject()
    {
        var (n, s) = Make();
        await n.NotifyAccessAddedAsync(Notice(), new[] { "newperson@x.com" }, "leader@x.com", CancellationToken.None);
        var msg = Assert.Single(s.Sent);
        Assert.Contains("You've been added", msg.Subject);
        Assert.Single(msg.To);
    }

    [Fact]
    public async Task Disabled_SendsNothing()
    {
        var (n, s) = Make(new NotificationOptions { Enabled = false });
        await n.NotifyCreatedAsync(Notice(), new[] { "pm@x.com" }, "leader@x.com", CancellationToken.None);
        Assert.Empty(s.Sent);
    }

    [Fact]
    public async Task SilentPractice_RedirectsToTestRecipient_AndTagsSubject()
    {
        var (n, s) = Make(new NotificationOptions
        {
            Enabled = true,
            BccAddress = "bcc@x.com",
            TestRecipient = "verifier@x.com",
            SilentPractices = new List<string> { "Marital Dissolution" },
        });

        var silent = Notice() with { Practice = "Marital Dissolution" };
        await n.NotifyCreatedAsync(silent, new[] { "pm@x.com", "team@x.com" }, "leader@x.com", CancellationToken.None);

        var msg = Assert.Single(s.Sent);
        Assert.Equal(new[] { "verifier@x.com" }, msg.To);      // redirected to the sole test recipient
        Assert.Null(msg.Bcc);                                   // bcc dropped on redirect
        Assert.StartsWith("[SILENT]", msg.Subject);
        Assert.Contains("intended: pm@x.com, team@x.com", msg.Subject);
    }

    [Fact]
    public async Task NonSilentPractice_EmailsRealRecipients_WhileAnotherIsSilent()
    {
        var (n, s) = Make(new NotificationOptions
        {
            Enabled = true,
            BccAddress = "bcc@x.com",
            TestRecipient = "verifier@x.com",
            SilentPractices = new List<string> { "Marital Dissolution" },
        });

        // Estate & Gift is NOT in the silent list -> real recipients, real bcc, no [SILENT] tag.
        await n.NotifyCreatedAsync(Notice(), new[] { "pm@x.com", "team@x.com" }, "leader@x.com", CancellationToken.None);

        var msg = Assert.Single(s.Sent);
        Assert.Equal(2, msg.To.Count);
        Assert.Contains("pm@x.com", msg.To);
        Assert.Equal("bcc@x.com", msg.Bcc);
        Assert.DoesNotContain("[SILENT]", msg.Subject);
    }

    [Fact]
    public async Task SilentPractice_MatchesMultiSelectToken()
    {
        var (n, s) = Make(new NotificationOptions
        {
            Enabled = true,
            TestRecipient = "verifier@x.com",
            SilentPractices = new List<string> { "Tangible Assets" },
        });

        // Multi-select practice value from a HubSpot deal.
        var multi = Notice() with { Practice = "Estate & Gift;Tangible Assets" };
        await n.NotifyCreatedAsync(multi, new[] { "pm@x.com" }, "leader@x.com", CancellationToken.None);

        var msg = Assert.Single(s.Sent);
        Assert.Equal(new[] { "verifier@x.com" }, msg.To);
        Assert.StartsWith("[SILENT]", msg.Subject);
    }

    [Fact]
    public void ScopingCreated_SubjectAndBody_ReflectScoping()
    {
        var scoping = Notice() with { Phase = WorkspacePhase.Scoping, IdLabel = "Opportunity #", IdValue = "PQ007180" };
        var (subject, html) = WorkspaceEmail.BuildCreated(scoping);
        Assert.Contains("New scoping dataroom", subject);
        Assert.Contains("PQ007180", subject);
        Assert.Contains("Open the dataroom", html);
        Assert.Contains("Client file-request link", html);   // upload link present -> button rendered
    }

    [Fact]
    public void Created_WithoutUploadLink_OmitsUploadButton()
    {
        var noLink = Notice() with { UploadLinkUrl = null };
        var (_, html) = WorkspaceEmail.BuildCreated(noLink);
        Assert.Contains("Open the dataroom", html);
        Assert.DoesNotContain("Client file-request link", html);
    }
}
