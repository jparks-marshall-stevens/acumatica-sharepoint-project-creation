using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Time.Testing;
using Moq;
using ProjectSync.Acumatica;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class ProjectSyncProcessorTests
{
    private static readonly DateTimeOffset Now = new(2026, 8, 5, 12, 0, 0, TimeSpan.Zero);

    private readonly Mock<IAcumaticaClient> _acumatica = new(MockBehavior.Strict);
    private readonly Mock<ISharePointDocumentSetService> _sharePoint = new(MockBehavior.Strict);
    private readonly Mock<ILastRunStore> _lastRun = new(MockBehavior.Strict);
    private readonly FakeTimeProvider _time = new(Now);
    private readonly StateOptions _state = new() { FirstRunLookbackHours = 24, OverlapMinutes = 5 };

    private ProjectSyncProcessor CreateSut() => new(
        _acumatica.Object,
        _sharePoint.Object,
        _lastRun.Object,
        Microsoft.Extensions.Options.Options.Create(_state),
        _time,
        NullLogger<ProjectSyncProcessor>.Instance);

    private static AcumaticaProject Project(string id, DateTimeOffset created) => new()
    {
        ProjectId = id,
        ProjectName = $"Name {id}",
        CustomerName = $"Customer {id}",
        ProjectManager = $"PM {id}",
        Practice = "Advisory",
        CreatedDateTime = created,
    };

    [Fact]
    public async Task FirstRun_QueriesWithLookback_AndCreatesSets()
    {
        DateTimeOffset? capturedFrom = null;
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync((DateTimeOffset?)null);

        var p1 = Project("P1", Now.AddHours(-2));
        var p2 = Project("P2", Now.AddHours(-1));
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Callback<DateTimeOffset, CancellationToken>((from, _) => capturedFrom = from)
            .ReturnsAsync(new[] { p1, p2 });

        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(It.IsAny<AcumaticaProject>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: true, "/url"));

        DateTimeOffset? savedWatermark = null;
        _lastRun.Setup(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Callback<DateTimeOffset, CancellationToken>((w, _) => savedWatermark = w)
            .Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(Now.AddHours(-24), capturedFrom); // first-run lookback
        Assert.Equal(2, result.Found);
        Assert.Equal(2, result.Created);
        Assert.Equal(0, result.Updated);
        Assert.Equal(p2.CreatedDateTime, savedWatermark); // newest processed
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(It.IsAny<AcumaticaProject>(), It.IsAny<CancellationToken>()), Times.Exactly(2));
    }

    [Fact]
    public async Task SubsequentRun_QueriesWithOverlap()
    {
        var lastRun = Now.AddMinutes(-30);
        DateTimeOffset? capturedFrom = null;
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(lastRun);
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Callback<DateTimeOffset, CancellationToken>((from, _) => capturedFrom = from)
            .ReturnsAsync(Array.Empty<AcumaticaProject>());

        await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(lastRun.AddMinutes(-5), capturedFrom); // overlap applied
        // No projects on a subsequent run must NOT overwrite the watermark.
        _lastRun.Verify(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task FirstRun_NoProjects_SetsWatermarkToNow()
    {
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync((DateTimeOffset?)null);
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(Array.Empty<AcumaticaProject>());
        _lastRun.Setup(s => s.SetLastRunAsync(Now, It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(0, result.Found);
        _lastRun.Verify(s => s.SetLastRunAsync(Now, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Failure_HaltsCycle_AndHoldsWatermarkBeforeFailedRecord()
    {
        var lastRun = Now.AddHours(-3);
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(lastRun);

        var p1 = Project("P1", Now.AddHours(-2));
        var p2 = Project("P2", Now.AddHours(-1)); // this one fails
        var p3 = Project("P3", Now.AddMinutes(-30));
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { p1, p2, p3 });

        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(p1, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: true, "/p1"));
        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(p2, It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("boom"));

        DateTimeOffset? savedWatermark = null;
        _lastRun.Setup(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Callback<DateTimeOffset, CancellationToken>((w, _) => savedWatermark = w)
            .Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.True(result.HadFailure);
        Assert.Equal(1, result.Created);
        Assert.Equal(p1.CreatedDateTime, savedWatermark); // stopped at P1; P2/P3 retried next cycle
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(p3, It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task CountsCreatedAndUpdatedSeparately()
    {
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(Now.AddHours(-2));

        var p1 = Project("P1", Now.AddMinutes(-50));
        var p2 = Project("P2", Now.AddMinutes(-40));
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { p1, p2 });

        _sharePoint.Setup(s => s.EnsureProjectDocumentSetAsync(p1, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: true, "/p1"));
        _sharePoint.Setup(s => s.EnsureProjectDocumentSetAsync(p2, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: false, "/p2")); // already existed
        _lastRun.Setup(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(1, result.Created);
        Assert.Equal(1, result.Updated);
    }
}
