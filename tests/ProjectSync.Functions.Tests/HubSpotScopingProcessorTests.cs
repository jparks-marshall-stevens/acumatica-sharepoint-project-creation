using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Time.Testing;
using Moq;
using ProjectSync.HubSpot;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class HubSpotScopingProcessorTests
{
    private static readonly DateTimeOffset Now = new(2026, 8, 21, 12, 0, 0, TimeSpan.Zero);

    private readonly Mock<IHubSpotClient> _hubspot = new(MockBehavior.Strict);
    private readonly Mock<ISharePointDocumentSetService> _sharePoint = new(MockBehavior.Strict);
    private readonly Mock<ILastRunStore> _store = new(MockBehavior.Strict);
    private readonly FakeTimeProvider _time = new(Now);
    private readonly HubSpotOptions _options = new();

    private HubSpotScopingProcessor CreateSut() => new(
        _hubspot.Object, _sharePoint.Object, _store.Object,
        Microsoft.Extensions.Options.Options.Create(_options), _time,
        NullLogger<HubSpotScopingProcessor>.Instance);

    [Fact]
    public async Task ResolvesDealOwner_ProjectManager_AndBothOriginators_OntoTheWorkspace()
    {
        _store.Setup(s => s.GetWatermarkAsync("hubspot-deals", It.IsAny<CancellationToken>()))
            .ReturnsAsync((DateTimeOffset?)null);
        _store.Setup(s => s.SetWatermarkAsync("hubspot-deals", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var deal = new HubSpotDeal
        {
            DealId = "D1",
            DealName = "Sample deal",
            Practice = "Marital Dissolution",
            OwnerId = "own", ProjectManagerId = "pm", OriginatorAId = "oa", OriginatorBId = "ob",
            CreatedAt = Now.AddHours(-2), ModifiedAt = Now.AddHours(-1),
        };
        _hubspot.Setup(h => h.GetDealsModifiedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { deal });
        _hubspot.Setup(h => h.ResolveCustomerNameAsync(deal, It.IsAny<CancellationToken>()))
            .ReturnsAsync("Sample Holdings LLC");
        _hubspot.Setup(h => h.GetOwnerEmailsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new Dictionary<string, string>
            {
                ["own"] = "owner@marshall-stevens.com",
                ["pm"] = "pm@marshall-stevens.com",
                ["oa"] = "orig.a@marshall-stevens.com",
                ["ob"] = "orig.b@marshall-stevens.com",
            });

        ScopingWorkspace? captured = null;
        _sharePoint.Setup(s => s.EnsureScopingWorkspaceAsync(It.IsAny<ScopingWorkspace>(), It.IsAny<CancellationToken>()))
            .Callback<ScopingWorkspace, CancellationToken>((w, _) => captured = w)
            .ReturnsAsync(new DocumentSetResult(Created: true, "/url"));

        await CreateSut().RunAsync(dryRun: false, CancellationToken.None);

        Assert.NotNull(captured);
        Assert.Equal("owner@marshall-stevens.com", captured!.OwnerEmail);
        Assert.Equal("pm@marshall-stevens.com", captured.ProjectManagerEmail);
        Assert.Equal("orig.a@marshall-stevens.com", captured.OriginatorAEmail);
        Assert.Equal("orig.b@marshall-stevens.com", captured.OriginatorBEmail);
    }

    [Fact]
    public async Task UnresolvableOrAbsentRoles_LeaveEmailsNull()
    {
        _store.Setup(s => s.GetWatermarkAsync("hubspot-deals", It.IsAny<CancellationToken>()))
            .ReturnsAsync((DateTimeOffset?)null);
        _store.Setup(s => s.SetWatermarkAsync("hubspot-deals", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Owner resolves; PM id isn't in the owner map; originators absent on the deal.
        var deal = new HubSpotDeal
        {
            DealId = "D2", Practice = "ESOP",
            OwnerId = "own", ProjectManagerId = "ghost",
            CreatedAt = Now.AddHours(-2), ModifiedAt = Now.AddHours(-1),
        };
        _hubspot.Setup(h => h.GetDealsModifiedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { deal });
        _hubspot.Setup(h => h.ResolveCustomerNameAsync(deal, It.IsAny<CancellationToken>()))
            .ReturnsAsync("Cust");
        _hubspot.Setup(h => h.GetOwnerEmailsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new Dictionary<string, string> { ["own"] = "owner@marshall-stevens.com" });

        ScopingWorkspace? captured = null;
        _sharePoint.Setup(s => s.EnsureScopingWorkspaceAsync(It.IsAny<ScopingWorkspace>(), It.IsAny<CancellationToken>()))
            .Callback<ScopingWorkspace, CancellationToken>((w, _) => captured = w)
            .ReturnsAsync(new DocumentSetResult(Created: true, "/url"));

        await CreateSut().RunAsync(dryRun: false, CancellationToken.None);

        Assert.NotNull(captured);
        Assert.Equal("owner@marshall-stevens.com", captured!.OwnerEmail);
        Assert.Null(captured.ProjectManagerEmail);   // id not in the owner map
        Assert.Null(captured.OriginatorAEmail);       // absent on the deal
        Assert.Null(captured.OriginatorBEmail);
    }
}
