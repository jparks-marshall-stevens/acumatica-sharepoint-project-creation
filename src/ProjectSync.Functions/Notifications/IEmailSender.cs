namespace ProjectSync.Notifications;

/// <summary>A composed email, ready to send.</summary>
public sealed record EmailMessage
{
    public required IReadOnlyList<string> To { get; init; }
    public string? Bcc { get; init; }
    public string? ReplyTo { get; init; }
    public required string Subject { get; init; }
    public required string HtmlBody { get; init; }
}

/// <summary>
/// Abstraction over "actually send the email." Swappable: a logging stub for now, a Microsoft Graph
/// implementation later — callers never change.
/// </summary>
public interface IEmailSender
{
    Task SendAsync(EmailMessage message, CancellationToken cancellationToken);
}
