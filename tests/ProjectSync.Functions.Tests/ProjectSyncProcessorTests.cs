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
    private readonly AcumaticaOptions _acumaticaOptions = new();

    private ProjectSyncProcessor CreateSut()
    {
        // Default: no team members unless a test overrides.
        _acumatica.Setup(a => a.GetTeamEmailsAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(Array.Empty<string>());
        return new(
            _acumatica.Object,
        _sharePoint.Object,
        _lastRun.Object,
        Microsoft.Extensions.Options.Options.Create(_state),
        Microsoft.Extensions.Options.Options.Create(_acumaticaOptions),
            _time,
            NullLogger<ProjectSyncProcessor>.Instance);
    }

    private static AcumaticaProject Project(string id, DateTimeOffset created, string practice = "Advisory") => new()
    {
        ProjectId = id,
        ProjectName = $"Name {id}",
        CustomerName = $"Customer {id}",
        ProjectManager = $"PM {id}",
        Practice = practice,
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
    public async Task PracticeAllowList_SkipsOthers_ButStillAdvancesWatermarkPastThem()
    {
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(Now.AddHours(-2));

        var gift = Project("P1", Now.AddMinutes(-50), practice: "estate & gift"); // case-insensitive match
        var litigation = Project("P2", Now.AddMinutes(-40), practice: "Commercial Litigation");
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { gift, litigation });

        // Only the Estate & Gift project should reach SharePoint.
        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(gift, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: true, "/p1"));

        DateTimeOffset? savedWatermark = null;
        _lastRun.Setup(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Callback<DateTimeOffset, CancellationToken>((w, _) => savedWatermark = w)
            .Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(1, result.Created);
        Assert.Equal(1, result.Skipped);
        Assert.Equal(litigation.CreatedDateTime, savedWatermark); // watermark advanced past the skipped one
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(litigation, It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ExcludedProjectId_IsSkipped_EvenWhenPracticeMatches()
    {
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _acumaticaOptions.ExcludedProjectIds = new List<string> { "X" };
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(Now.AddHours(-2));

        // "X" has the right practice but is an internal moniker; must never reach SharePoint.
        var internalProj = Project("X", Now.AddMinutes(-50), practice: "Estate & Gift");
        var realProj = Project("10-31-21-74661", Now.AddMinutes(-40), practice: "Estate & Gift");
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { internalProj, realProj });

        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(realProj, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: true, "/real"));
        _lastRun.Setup(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var result = await CreateSut().RunAsync(CancellationToken.None);

        Assert.Equal(1, result.Created);
        Assert.Equal(1, result.Skipped);
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(internalProj, It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task DryRun_PlansButCreatesNothing_AndDoesNotPersistWatermark()
    {
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync((DateTimeOffset?)null);

        var gift = Project("10-31-21-74661", Now.AddMinutes(-40), practice: "Estate & Gift");
        var other = Project("P2", Now.AddMinutes(-30), practice: "Commercial Litigation");
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { gift, other });

        _sharePoint
            .Setup(s => s.PlanDocumentSet(gift))
            .Returns(new DocumentSetPlan("https://site", "Project Documents", "Estate & Gift", "10-31-21-74661"));

        var result = await CreateSut().RunAsync(new RunOptions { DryRun = true }, CancellationToken.None);

        Assert.True(result.DryRun);
        Assert.Equal(1, result.Planned);
        Assert.Equal(0, result.Created);
        Assert.Equal(1, result.Skipped);
        Assert.NotNull(result.Plan);
        Assert.Single(result.Plan!);
        Assert.Equal("Estate & Gift", result.Plan![0].TargetFolder);

        // Dry run must not create anything or move the watermark.
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(It.IsAny<AcumaticaProject>(), It.IsAny<CancellationToken>()), Times.Never);
        _lastRun.Verify(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task TargetedRun_ProcessesOnlyThatProject_AndDoesNotPersistWatermark()
    {
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _lastRun.Setup(s => s.GetLastRunAsync(It.IsAny<CancellationToken>())).ReturnsAsync(Now.AddHours(-2));

        var target = Project("10-31-21-74663", Now.AddMinutes(-50), practice: "Estate & Gift");
        var other = Project("10-31-21-74664", Now.AddMinutes(-40), practice: "Estate & Gift");
        _acumatica
            .Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { target, other });

        _sharePoint
            .Setup(s => s.EnsureProjectDocumentSetAsync(target, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DocumentSetResult(Created: false, "/updated"));

        var result = await CreateSut().RunAsync(
            new RunOptions { OnlyProjectId = "10-31-21-74663" }, CancellationToken.None);

        Assert.Equal(0, result.Created);
        Assert.Equal(1, result.Updated);
        _sharePoint.Verify(s => s.EnsureProjectDocumentSetAsync(other, It.IsAny<CancellationToken>()), Times.Never);
        // Targeted runs must not move the watermark.
        _lastRun.Verify(s => s.SetLastRunAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Never);
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

    [Fact]
    public async Task ReconcileIncremental_NoTeamChanges_ShortCircuits_NoSharePoint()
    {
        var wm = Now.AddMinutes(-30);
        _acumatica.Setup(a => a.GetTeamRowsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { new TeamMemberRow("P1", "a@x.com", Now.AddHours(-1)) }); // older than watermark
        _lastRun.Setup(s => s.GetWatermarkAsync("reconcile-team", It.IsAny<CancellationToken>())).ReturnsAsync(wm);

        var result = await CreateSut().ReconcileIncrementalAsync(CancellationToken.None);

        Assert.Equal(0, result.Updated);
        _sharePoint.Verify(s => s.ReconcileAsync(It.IsAny<IReadOnlyList<AcumaticaProject>>(), It.IsAny<IReadOnlySet<string>>(), It.IsAny<CancellationToken>()), Times.Never);
        _acumatica.Verify(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ReconcileIncremental_TeamChanged_ReconcilesOnlyChangedIds()
    {
        var wm = Now.AddMinutes(-30);
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _acumatica.Setup(a => a.GetTeamRowsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[]
            {
                new TeamMemberRow("P1", "a@x.com", Now.AddMinutes(-5)),  // changed
                new TeamMemberRow("P2", "b@x.com", Now.AddHours(-2)),    // unchanged
            });
        _lastRun.Setup(s => s.GetWatermarkAsync("reconcile-team", It.IsAny<CancellationToken>())).ReturnsAsync(wm);
        _lastRun.Setup(s => s.SetWatermarkAsync("reconcile-team", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);
        _acumatica.Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { Project("P1", Now, "Estate & Gift"), Project("P2", Now, "Estate & Gift") });

        IReadOnlySet<string>? capturedIds = null;
        _sharePoint.Setup(s => s.ReconcileAsync(It.IsAny<IReadOnlyList<AcumaticaProject>>(), It.IsAny<IReadOnlySet<string>>(), It.IsAny<CancellationToken>()))
            .Callback<IReadOnlyList<AcumaticaProject>, IReadOnlySet<string>?, CancellationToken>((_, ids, _) => capturedIds = ids)
            .ReturnsAsync(new ReconcileResult { Updated = 1 });

        await CreateSut().ReconcileIncrementalAsync(CancellationToken.None);

        Assert.NotNull(capturedIds);
        Assert.Contains("P1", capturedIds!);
        Assert.DoesNotContain("P2", capturedIds!);
        _lastRun.Verify(s => s.SetWatermarkAsync("reconcile-team", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ReconcileFull_PassesNullIds_ForAllTracked()
    {
        _acumaticaOptions.IncludedPractices = new List<string> { "Estate & Gift" };
        _acumatica.Setup(a => a.GetTeamRowsAsync(It.IsAny<CancellationToken>())).ReturnsAsync(Array.Empty<TeamMemberRow>());
        _acumatica.Setup(a => a.GetProjectsCreatedAfterAsync(It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { Project("P1", Now, "Estate & Gift") });
        _lastRun.Setup(s => s.SetWatermarkAsync("reconcile-team", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var sawNull = false;
        _sharePoint.Setup(s => s.ReconcileAsync(It.IsAny<IReadOnlyList<AcumaticaProject>>(), It.IsAny<IReadOnlySet<string>>(), It.IsAny<CancellationToken>()))
            .Callback<IReadOnlyList<AcumaticaProject>, IReadOnlySet<string>?, CancellationToken>((_, ids, _) => sawNull = ids is null)
            .ReturnsAsync(new ReconcileResult { Considered = 1, Updated = 1 });

        await CreateSut().ReconcileFullAsync(CancellationToken.None);

        Assert.True(sawNull); // full sweep reconciles all tracked (onlyProjectIds == null)
    }
}
