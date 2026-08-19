using Microsoft.Extensions.Logging;

namespace ProjectSync.Notifications;

/// <summary>
/// Stub sender: it does NOT send. It logs the fully-composed message (recipients, subject, and body)
/// so the notification flow can be built and reviewed before a real mail transport is wired up. Swap the
/// DI registration for a Graph-backed sender to go live — nothing else changes.
/// </summary>
public sealed class LoggingEmailSender : IEmailSender
{
    private readonly ILogger<LoggingEmailSender> _logger;

    public LoggingEmailSender(ILogger<LoggingEmailSender> logger) => _logger = logger;

    public Task SendAsync(EmailMessage message, CancellationToken cancellationToken)
    {
        _logger.LogInformation(
            "[EMAIL STUB — not sent] To: {To} | Bcc: {Bcc} | Subject: {Subject}\n{Body}",
            string.Join(", ", message.To),
            string.IsNullOrWhiteSpace(message.Bcc) ? "(none)" : message.Bcc,
            message.Subject,
            message.HtmlBody);
        return Task.CompletedTask;
    }
}
