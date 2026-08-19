using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using ProjectSync.Options;
using ProjectSync.SharePoint;

namespace ProjectSync.Notifications;

/// <summary>
/// Sends email via Microsoft Graph <c>sendMail</c> as the configured no-reply mailbox, using the same
/// app-only certificate the SharePoint path uses. The app has Graph <c>Mail.Send</c>, scoped by an
/// Exchange Application Access Policy to only the no-reply mailbox — so it can send as that address and
/// nothing else. Throws on failure so the caller (the notifier) can log it; the notifier is fail-soft, so
/// a send failure never disrupts the sync.
/// </summary>
public sealed class GraphEmailSender : IEmailSender
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };

    private readonly SharePointOptions _sp;
    private readonly NotificationOptions _options;
    private readonly ILogger<GraphEmailSender> _logger;
    private readonly Lazy<IConfidentialClientApplication> _app;

    public GraphEmailSender(
        SharePointContextFactory contextFactory,
        IOptions<SharePointOptions> sharePointOptions,
        IOptions<NotificationOptions> options,
        ILogger<GraphEmailSender> logger)
    {
        _sp = sharePointOptions.Value;
        _options = options.Value;
        _logger = logger;
        _app = new Lazy<IConfidentialClientApplication>(() =>
            ConfidentialClientApplicationBuilder.Create(_sp.ClientId)
                .WithCertificate(contextFactory.Certificate)
                .WithTenantId(_sp.AzureAdTenant)
                .Build());
    }

    public async Task SendAsync(EmailMessage message, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.FromAddress))
        {
            _logger.LogError("Notifications:FromAddress is not set; cannot send '{Subject}'.", message.Subject);
            return;
        }

        var token = (await _app.Value
            .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
            .ExecuteAsync(cancellationToken)).AccessToken;

        var msg = new Dictionary<string, object?>
        {
            ["subject"] = message.Subject,
            ["body"] = new Dictionary<string, object?> { ["contentType"] = "HTML", ["content"] = message.HtmlBody },
            ["toRecipients"] = message.To.Select(Recipient).ToArray(),
        };
        if (!string.IsNullOrWhiteSpace(message.Bcc))
        {
            msg["bccRecipients"] = new[] { Recipient(message.Bcc!) };
        }
        if (!string.IsNullOrWhiteSpace(message.ReplyTo))
        {
            msg["replyTo"] = new[] { Recipient(message.ReplyTo!) };
        }

        var payload = new Dictionary<string, object?> { ["message"] = msg, ["saveToSentItems"] = false };
        var url = $"{GraphBase}/users/{Uri.EscapeDataString(_options.FromAddress)}/sendMail";

        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"),
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using var response = await Http.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            _logger.LogError("Graph sendMail failed ({Status}) as {From}: {Body}",
                (int)response.StatusCode, _options.FromAddress, body.Length > 500 ? body[..500] : body);
            throw new HttpRequestException($"Graph sendMail failed ({(int)response.StatusCode}).");
        }

        _logger.LogInformation("Sent '{Subject}' to {To} as {From}.",
            message.Subject, string.Join(", ", message.To), _options.FromAddress);
    }

    private static Dictionary<string, object?> Recipient(string address) =>
        new() { ["emailAddress"] = new Dictionary<string, object?> { ["address"] = address } };
}
