using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ProjectSync.Options;

namespace ProjectSync.Notifications;

/// <summary>
/// Turns workspace events into emails: computes the recipient set (excluding the practice leader),
/// builds the message, and hands it to the <see cref="IEmailSender"/>. Entirely fail-soft — a
/// notification failure is logged and never disrupts the sync that triggered it.
/// </summary>
public sealed class WorkspaceNotifier
{
    private readonly IEmailSender _sender;
    private readonly NotificationOptions _options;
    private readonly ILogger<WorkspaceNotifier> _logger;

    public WorkspaceNotifier(
        IEmailSender sender,
        IOptions<NotificationOptions> options,
        ILogger<WorkspaceNotifier> logger)
    {
        _sender = sender;
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>Notifies everyone with access (minus the practice leader) that a workspace was created.</summary>
    public Task NotifyCreatedAsync(
        WorkspaceNotice notice, IEnumerable<string?> recipients, string? leaderEmail, CancellationToken cancellationToken)
        => SendAsync(() => WorkspaceEmail.BuildCreated(notice), recipients, leaderEmail, "created", notice, cancellationToken);

    /// <summary>Notifies people who were just granted access (minus the practice leader) that they can start.</summary>
    public Task NotifyAccessAddedAsync(
        WorkspaceNotice notice, IEnumerable<string?> newlyAdded, string? leaderEmail, CancellationToken cancellationToken)
        => SendAsync(() => WorkspaceEmail.BuildAccessAdded(notice), newlyAdded, leaderEmail, "access-added", notice, cancellationToken);

    private async Task SendAsync(
        Func<(string Subject, string Html)> build,
        IEnumerable<string?> recipients,
        string? leaderEmail,
        string kind,
        WorkspaceNotice notice,
        CancellationToken cancellationToken)
    {
        if (!_options.Enabled)
        {
            return;
        }

        try
        {
            var leader = leaderEmail?.Trim();
            var to = recipients
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Select(r => r!.Trim())
                .Where(r => string.IsNullOrEmpty(leader) || !r.Equals(leader, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (to.Count == 0)
            {
                _logger.LogInformation("No {Kind} notification for {Customer}: no recipients after exclusions.", kind, notice.CustomerName);
                return;
            }

            var (subject, html) = build();
            await _sender.SendAsync(new EmailMessage
            {
                To = to,
                Bcc = string.IsNullOrWhiteSpace(_options.BccAddress) ? null : _options.BccAddress.Trim(),
                Subject = subject,
                HtmlBody = html,
            }, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to send {Kind} notification for {Customer}; continuing.", kind, notice.CustomerName);
        }
    }
}
