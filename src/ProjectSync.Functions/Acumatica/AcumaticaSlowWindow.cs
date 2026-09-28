using System.Globalization;
using ProjectSync.Options;

namespace ProjectSync.Acumatica;

/// <summary>
/// Recognises the daily Acumatica stall (see <see cref="AcumaticaOptions.SlowWindowStartUtc"/>) so the
/// timer functions can skip the run quietly instead of failing and paging.
/// </summary>
public static class AcumaticaSlowWindow
{
    /// <summary>
    /// True when <paramref name="ex"/> is an HttpClient timeout (not a host-shutdown cancellation) and the
    /// run started inside the configured daily window.
    /// </summary>
    public static bool IsExpectedTimeout(Exception ex, DateTimeOffset runStartedUtc, AcumaticaOptions options,
        CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested || !IsHttpTimeout(ex))
        {
            return false;
        }

        return IsInWindow(runStartedUtc, options);
    }

    public static bool IsInWindow(DateTimeOffset utc, AcumaticaOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.SlowWindowStartUtc) || options.SlowWindowMinutes <= 0
            || !TimeSpan.TryParseExact(options.SlowWindowStartUtc.Trim(), @"hh\:mm", CultureInfo.InvariantCulture,
                out var start))
        {
            return false;
        }

        // Compare time-of-day on a circular clock so a window crossing midnight still works.
        var sinceStart = (utc.UtcDateTime.TimeOfDay - start + TimeSpan.FromDays(1)).Ticks % TimeSpan.TicksPerDay;
        return sinceStart < TimeSpan.FromMinutes(options.SlowWindowMinutes).Ticks;
    }

    // HttpClient.Timeout surfaces as TaskCanceledException with an inner TimeoutException.
    private static bool IsHttpTimeout(Exception? ex)
    {
        for (; ex is not null; ex = ex.InnerException)
        {
            if (ex is TimeoutException)
            {
                return true;
            }
        }

        return false;
    }
}
