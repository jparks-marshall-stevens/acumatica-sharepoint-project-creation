using ProjectSync.Acumatica;
using ProjectSync.Options;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class AcumaticaSlowWindowTests
{
    private static readonly AcumaticaOptions Options = new() { SlowWindowStartUtc = "10:00", SlowWindowMinutes = 25 };

    // What HttpClient throws when HttpClient.Timeout elapses.
    private static Exception HttpTimeout() => new TaskCanceledException(
        "The request was canceled due to the configured HttpClient.Timeout of 100 seconds elapsing.",
        new TimeoutException("The operation was canceled."));

    private static DateTimeOffset Utc(int hour, int minute) => new(2026, 9, 28, hour, minute, 0, TimeSpan.Zero);

    [Theory]
    [InlineData(10, 0)]
    [InlineData(10, 15)]
    [InlineData(10, 24)]
    public void HttpTimeout_InsideWindow_IsExpected(int hour, int minute) =>
        Assert.True(AcumaticaSlowWindow.IsExpectedTimeout(HttpTimeout(), Utc(hour, minute), Options, CancellationToken.None));

    [Theory]
    [InlineData(9, 45)]
    [InlineData(10, 30)]
    [InlineData(14, 0)]
    public void HttpTimeout_OutsideWindow_IsNotExpected(int hour, int minute) =>
        Assert.False(AcumaticaSlowWindow.IsExpectedTimeout(HttpTimeout(), Utc(hour, minute), Options, CancellationToken.None));

    [Fact]
    public void OtherErrors_InsideWindow_AreNotExpected()
    {
        Assert.False(AcumaticaSlowWindow.IsExpectedTimeout(
            new InvalidOperationException("Acumatica token request failed (400): invalid_grant"),
            Utc(10, 0), Options, CancellationToken.None));
        Assert.False(AcumaticaSlowWindow.IsExpectedTimeout(
            new TaskCanceledException("The operation was canceled."), Utc(10, 0), Options, CancellationToken.None));
    }

    [Fact]
    public void HostShutdown_IsNotExpected()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        Assert.False(AcumaticaSlowWindow.IsExpectedTimeout(HttpTimeout(), Utc(10, 0), Options, cts.Token));
    }

    [Fact]
    public void BlankStart_DisablesWindow() =>
        Assert.False(AcumaticaSlowWindow.IsExpectedTimeout(HttpTimeout(), Utc(10, 0),
            new AcumaticaOptions { SlowWindowStartUtc = "" }, CancellationToken.None));

    [Fact]
    public void WindowCrossingMidnight_Works()
    {
        var options = new AcumaticaOptions { SlowWindowStartUtc = "23:50", SlowWindowMinutes = 20 };
        Assert.True(AcumaticaSlowWindow.IsInWindow(Utc(23, 55), options));
        Assert.True(AcumaticaSlowWindow.IsInWindow(Utc(0, 5), options));
        Assert.False(AcumaticaSlowWindow.IsInWindow(Utc(0, 15), options));
    }
}
