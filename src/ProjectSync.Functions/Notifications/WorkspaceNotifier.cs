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
        => SendAsync(() => WorkspaceEmail.BuildCreated(notice, _options.LogoUrl), recipients, leaderEmail, "created", notice, cancellationToken);

    /// <summary>Notifies people who were just granted access (minus the practice leader) that they can start.</summary>
    public Task NotifyAccessAddedAsync(
        WorkspaceNotice notice, IEnumerable<string?> newlyAdded, string? leaderEmail, CancellationToken cancellationToken)
        => SendAsync(() => WorkspaceEmail.BuildAccessAdded(notice, _options.LogoUrl), newlyAdded, leaderEmail, "access-added", notice, cancellationToken);

    /// <summary>
    /// Notifies everyone with access to a workspace that the client uploaded files. <paramref name="excludeEmail"/>
    /// is the practice leader to drop for engagements (Bruce isn't delivering execution work); pass null for
    /// scoping rooms so he stays included.
    /// </summary>
    public Task NotifyClientUploadAsync(
        WorkspaceNotice notice, IReadOnlyList<string> fileNames, string uploadsFolderUrl,
        IEnumerable<string?> recipients, string? excludeEmail, CancellationToken cancellationToken)
        => SendAsync(() => WorkspaceEmail.BuildClientUpload(notice, fileNames, uploadsFolderUrl, _options.LogoUrl),
            recipients, excludeEmail, "client-upload", notice, cancellationToken);

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

            // Redirect the whole thing to the test recipient so no live person is emailed, and surface the
            // intended recipients in the subject. Triggered globally by TestMode, or per-practice by
            // SilentPractices (switch a new practice on for verification without diverting a live one).
            var actualTo = to;
            string? bcc = string.IsNullOrWhiteSpace(_options.BccAddress) ? null : _options.BccAddress.Trim();
            var redirectReason = _options.TestMode ? "TEST" : (IsSilentPractice(notice.Practice) ? "SILENT" : null);
            if (redirectReason is not null)
            {
                if (string.IsNullOrWhiteSpace(_options.TestRecipient))
                {
                    _logger.LogWarning("{Reason} redirect is on but TestRecipient is blank; not sending {Kind} for {Customer}.", redirectReason, kind, notice.CustomerName);
                    return;
                }

                subject = $"[{redirectReason}] {subject}  (intended: {string.Join(", ", to)})";
                actualTo = new List<string> { _options.TestRecipient.Trim() };
                bcc = null;
            }

            await _sender.SendAsync(new EmailMessage
            {
                To = actualTo,
                Bcc = bcc,
                ReplyTo = string.IsNullOrWhiteSpace(_options.ReplyToAddress) ? null : _options.ReplyToAddress.Trim(),
                Subject = subject,
                HtmlBody = html,
            }, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to send {Kind} notification for {Customer}; continuing.", kind, notice.CustomerName);
        }
    }

    /// <summary>
    /// True when the workspace's practice is in <see cref="NotificationOptions.SilentPractices"/> — its
    /// emails should be redirected to the test recipient. A HubSpot practice can be a ';'-delimited
    /// multi-select, so match if the whole value or ANY of its tokens matches a configured entry.
    /// </summary>
    private bool IsSilentPractice(string? practice)
    {
        if (_options.SilentPractices is not { Count: > 0 } || string.IsNullOrWhiteSpace(practice))
        {
            return false;
        }

        var tokens = practice.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return _options.SilentPractices.Any(sp =>
            !string.IsNullOrWhiteSpace(sp) &&
            (string.Equals(sp.Trim(), practice.Trim(), StringComparison.OrdinalIgnoreCase) ||
             tokens.Any(t => string.Equals(sp.Trim(), t, StringComparison.OrdinalIgnoreCase))));
    }
}
