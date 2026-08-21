using System.Net;
using System.Text;

namespace ProjectSync.Notifications;

/// <summary>The lifecycle phase a workspace notification is about.</summary>
public enum WorkspacePhase
{
    Scoping,
    Execution,
}

/// <summary>Everything an email needs to describe a workspace to its recipients.</summary>
public sealed record WorkspaceNotice
{
    public required WorkspacePhase Phase { get; init; }
    public required string CustomerName { get; init; }

    /// <summary>Project name (execution) or deal/engagement name (scoping).</summary>
    public string? EngagementName { get; init; }

    /// <summary>Label for the identifier row, e.g. "Project ID" or "Opportunity #".</summary>
    public string? IdLabel { get; init; }
    public string? IdValue { get; init; }

    public string? ProjectManager { get; init; }
    public string? Practice { get; init; }

    /// <summary>Absolute URL to the document set (the "dataroom").</summary>
    public required string DataroomUrl { get; init; }

    /// <summary>Absolute URL of the client file-request (upload) link, if one exists.</summary>
    public string? UploadLinkUrl { get; init; }
}

/// <summary>Builds the subject line and HTML body for each workspace email (email-safe, inline styles).</summary>
public static class WorkspaceEmail
{
    // Marshall & Stevens brand teal (deep enough for legible white text/logo on the header bar).
    private const string BrandTeal = "#2C7E85";

    // Cap the number of filenames listed inline before summarizing the remainder.
    private const int MaxFilesListed = 25;

    public static (string Subject, string Html) BuildCreated(WorkspaceNotice n, string? logoUrl = null)
    {
        var kind = n.Phase == WorkspacePhase.Scoping ? "scoping" : "project";
        var subject = $"New {kind} dataroom — {Suffix(n)}";
        var kicker = Kicker(n.Phase == WorkspacePhase.Scoping ? "Scoping" : "Active project", n.Practice);
        var title = n.Phase == WorkspacePhase.Scoping ? "A scoping dataroom is ready" : "A project dataroom is ready";
        const string intro = "A workspace has been created for the engagement below. You're receiving this because you have access to it.";

        var inner = new StringBuilder();
        inner.Append(Para(intro));
        inner.Append(DetailTable(n));
        inner.Append(Buttons(
            (label: "Open the dataroom", url: n.DataroomUrl, primary: true),
            (label: "Client file-request link", url: n.UploadLinkUrl ?? string.Empty, primary: false)));
        if (!string.IsNullOrWhiteSpace(n.UploadLinkUrl))
        {
            inner.Append(UploadNote());
        }

        return (subject, Shell(BrandTeal, kicker, title, logoUrl, inner.ToString()));
    }

    public static (string Subject, string Html) BuildAccessAdded(WorkspaceNotice n, string? logoUrl = null)
    {
        var subject = $"You've been added — {Suffix(n)}";
        var kicker = Kicker("Access granted", n.Practice);
        const string title = "You now have access to a dataroom";
        const string intro = "You've been given access to the workspace below, so you can start work. You're receiving this because your access was just added.";

        var inner = new StringBuilder();
        inner.Append(Para(intro));
        inner.Append(DetailTable(n));
        inner.Append(Buttons(
            (label: "Open the dataroom", url: n.DataroomUrl, primary: true),
            (label: "Client file-request link", url: n.UploadLinkUrl ?? string.Empty, primary: false)));
        if (!string.IsNullOrWhiteSpace(n.UploadLinkUrl))
        {
            inner.Append(UploadNote());
        }

        return (subject, Shell(BrandTeal, kicker, title, logoUrl, inner.ToString()));
    }

    /// <summary>Email for "the client uploaded files" — states how many and lists their names.</summary>
    public static (string Subject, string Html) BuildClientUpload(
        WorkspaceNotice n, IReadOnlyList<string> fileNames, string uploadsFolderUrl, string? logoUrl = null)
    {
        var count = fileNames.Count;
        var noun = count == 1 ? "file" : "files";
        var subject = $"Client uploaded {count} {noun} — {Suffix(n)}";
        var kicker = Kicker("Client upload", n.Practice);
        var title = $"{count} new client {noun} uploaded";
        var intro = $"The client uploaded {count} new {noun} to the client uploads folder for the engagement below. You're receiving this because you have access to it.";

        var inner = new StringBuilder();
        inner.Append(Para(intro));
        inner.Append(DetailTable(n));
        inner.Append(FileList(fileNames));
        inner.Append(Buttons(
            (label: "Open the client uploads", url: uploadsFolderUrl, primary: true)));

        return (subject, Shell(BrandTeal, kicker, title, logoUrl, inner.ToString()));
    }

    private static string Suffix(WorkspaceNotice n) =>
        string.IsNullOrWhiteSpace(n.IdValue) ? n.CustomerName : $"{n.CustomerName} ({n.IdValue})";

    /// <summary>
    /// Header kicker "{status} &middot; {PRACTICE}" (the practice is uppercased by CSS in the shell). Reads
    /// the workspace's actual practice rather than hard-coding one; falls back to just the status when the
    /// practice is blank, and shows only the first value of a ';'-delimited multi-select. HTML-encoded.
    /// </summary>
    private static string Kicker(string status, string? practice)
    {
        if (string.IsNullOrWhiteSpace(practice))
        {
            return status;
        }

        var first = practice
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault() ?? practice.Trim();
        return $"{status} · {Html(first)}";
    }

    // ----- HTML fragments -----

    private static string Para(string text) =>
        $"<p style=\"margin:0 0 16px;color:#354039;\">{Html(text)}</p>";

    private static string DetailTable(WorkspaceNotice n)
    {
        var rows = new StringBuilder();
        Row(rows, "Client", n.CustomerName);
        Row(rows, n.Phase == WorkspacePhase.Scoping ? "Engagement" : "Project", n.EngagementName);
        Row(rows, n.IdLabel, n.IdValue);
        Row(rows, "Project manager", n.ProjectManager);
        return
            "<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse:collapse;margin:4px 0 22px;background:#f6f8f5;border:1px solid #e2e7de;border-radius:8px;overflow:hidden;\">" +
            rows + "</table>";
    }

    private static void Row(StringBuilder sb, string? label, string? value)
    {
        if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        sb.Append("<tr><td style=\"padding:9px 14px;font-size:11px;letter-spacing:.05em;text-transform:uppercase;font-weight:700;color:#6a746a;width:40%;border-bottom:1px solid #e8ede4;\">")
          .Append(Html(label))
          .Append("</td><td style=\"padding:9px 14px;font-size:14px;font-weight:600;color:#1f2a24;border-bottom:1px solid #e8ede4;\">")
          .Append(Html(value))
          .Append("</td></tr>");
    }

    private static string FileList(IReadOnlyList<string> fileNames)
    {
        var shown = fileNames.Take(MaxFilesListed).ToList();
        var items = new StringBuilder();
        foreach (var f in shown)
        {
            items.Append("<li style=\"margin:0 0 4px;\">").Append(Html(f)).Append("</li>");
        }

        var more = fileNames.Count > shown.Count
            ? $"<p style=\"margin:8px 0 0;font-size:13px;color:#6a746a;\">+{fileNames.Count - shown.Count} more…</p>"
            : string.Empty;

        return
            "<div style=\"margin:0 0 22px;padding:14px 16px;background:#f6f8f5;border:1px solid #e2e7de;border-radius:8px;\">" +
            "<div style=\"font-size:11px;letter-spacing:.05em;text-transform:uppercase;font-weight:700;color:#6a746a;margin:0 0 8px;\">Files uploaded</div>" +
            "<ul style=\"margin:0;padding:0 0 0 18px;font-size:14px;color:#1f2a24;\">" + items + "</ul>" + more +
            "</div>";
    }

    private static string Buttons(params (string label, string url, bool primary)[] buttons)
    {
        var sb = new StringBuilder("<div style=\"margin:6px 0 4px;\">");
        foreach (var b in buttons)
        {
            if (string.IsNullOrWhiteSpace(b.url))
            {
                continue;
            }

            var style = b.primary
                ? $"background:{BrandTeal};color:#ffffff;"
                : $"background:#ffffff;color:{BrandTeal};border:1.5px solid {BrandTeal};";
            sb.Append($"<a href=\"{Attr(b.url)}\" style=\"display:inline-block;text-decoration:none;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:12px 22px;border-radius:7px;margin:0 8px 10px 0;{style}\">{Html(b.label)}</a>");
        }

        return sb.Append("</div>").ToString();
    }

    private static string UploadNote() =>
        "<p style=\"font-size:13px;color:#5b655c;background:#f2f5ef;border-left:3px solid #cbd3c6;padding:10px 14px;border-radius:0 6px 6px 0;margin:4px 0 0;\">Send the <strong>client file-request link</strong> to the client to collect their documents. They can upload without signing in, and they can't see anything already in the folder.</p>";

    private static string Shell(string accent, string kicker, string title, string? logoUrl, string innerHtml)
    {
        var logo = string.IsNullOrWhiteSpace(logoUrl)
            ? string.Empty
            : $"<img src=\"{Attr(logoUrl!)}\" alt=\"Marshall &amp; Stevens\" width=\"150\" style=\"display:block;height:28px;width:auto;margin:0 0 14px;border:0;\" />";

        return $@"<div style=""background:#eef1ec;padding:24px 0;font-family:Arial,Helvetica,sans-serif;"">
<table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0""><tr><td align=""center"">
<table role=""presentation"" width=""600"" cellpadding=""0"" cellspacing=""0"" style=""max-width:600px;background:#ffffff;border-radius:10px;overflow:hidden;"">
  <tr><td style=""padding:22px 28px 18px;background:#ffffff;border-bottom:3px solid {accent};"">
    {logo}<div style=""font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:{accent};margin:0 0 4px;"">{kicker}</div>
    <div style=""font-family:Georgia,'Times New Roman',serif;font-size:21px;line-height:1.2;color:#1f2a24;"">{title}</div>
  </td></tr>
  <tr><td style=""padding:24px 28px 28px;color:#1f2a24;font-size:15px;line-height:1.62;"">
    {innerHtml}
  </td></tr>
  <tr><td style=""padding:16px 28px 22px;border-top:1px solid #ecefe8;font-size:11.5px;line-height:1.5;color:#8a938a;"">
    Automated message from the Marshall &amp; Stevens project workspace sync. You're receiving it because you have access to this dataroom. Please don't reply to this address.
  </td></tr>
</table>
</td></tr></table>
</div>";
    }

    private static string Html(string s) => WebUtility.HtmlEncode(s);

    private static string Attr(string s) => WebUtility.HtmlEncode(s);
}
