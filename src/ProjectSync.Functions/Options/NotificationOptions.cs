namespace ProjectSync.Options;

/// <summary>
/// Settings for workspace notification emails. Sending is pluggable: today an <c>IEmailSender</c> stub
/// logs the composed message; a real Microsoft Graph sender is wired later, once the sending mailbox and
/// a scoped <c>Mail.Send</c> grant are in place.
/// </summary>
public sealed class NotificationOptions
{
    public const string SectionName = "Notifications";

    /// <summary>Master switch. When false, nothing is composed or sent.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// Mailbox the emails are sent from (a dedicated/shared mailbox — NOT a person). Left blank until the
    /// mailbox is chosen; the stub sender logs regardless, so the flow is testable before this is set.
    /// </summary>
    public string FromAddress { get; set; } = string.Empty;

    /// <summary>Friendly From name shown to recipients.</summary>
    public string FromDisplayName { get; set; } = "Marshall & Stevens Projects";

    /// <summary>Where replies go. Blank = no reply-to header.</summary>
    public string ReplyToAddress { get; set; } = string.Empty;

    /// <summary>
    /// BCC'd on every notification (for oversight during rollout). Blank = no BCC.
    /// </summary>
    public string BccAddress { get; set; } = string.Empty;

    /// <summary>
    /// Absolute URL of the Marshall &amp; Stevens logo shown in the email header (white wordmark on the
    /// brand-teal bar). Blank = no logo. NOTE: many email clients don't render SVG; host a PNG and point
    /// this at it for the broadest support.
    /// </summary>
    public string LogoUrl { get; set; } = string.Empty;

    /// <summary>
    /// Safety valve for rollout: when true, EVERY notification is redirected to <see cref="TestRecipient"/>
    /// only — no live recipients receive anything. The intended recipients are shown in the subject so the
    /// tester can see who it would have gone to. Turn off to go live.
    /// </summary>
    public bool TestMode { get; set; } = false;

    /// <summary>The sole recipient while <see cref="TestMode"/> is on.</summary>
    public string TestRecipient { get; set; } = string.Empty;
}
