using System.Security.Cryptography;
using System.Text;
using ProjectSync.Acumatica;

namespace ProjectSync.SharePoint;

/// <summary>
/// Computes a stable hash of everything the sync writes for a project (metadata + the set of people
/// granted access). The reconcile compares this against the value last stamped on the document set;
/// if they match, nothing changed and the set is skipped — so the daily full sweep only writes deltas.
/// </summary>
public static class ReconcileSignature
{
    // Separator that won't appear in names/emails.
    private const string Delimiter = "~|~";

    /// <summary>The grantee emails (PM + practice leader + practice admins + team), normalized and sorted.</summary>
    public static IReadOnlyList<string> GranteeEmails(AcumaticaProject project, string? leaderEmail, IEnumerable<string>? adminEmails = null)
    {
        var pm = string.IsNullOrWhiteSpace(project.ProjectManagerEmail) ? project.ProjectManager : project.ProjectManagerEmail;
        var all = new List<string?> { pm, project.PrincipalInChargeEmail, leaderEmail };
        all.AddRange(project.TeamEmails);
        if (adminEmails is not null) all.AddRange(adminEmails);
        return all
            .Where(e => !string.IsNullOrWhiteSpace(e))
            .Select(e => e!.Trim().ToLowerInvariant())
            .Distinct()
            .OrderBy(e => e, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>Hash of the metadata + grantee set (hex; fits a single-line text column).</summary>
    public static string Compute(AcumaticaProject project, string? leaderEmail, IEnumerable<string>? adminEmails = null,
        string? membersPermissionLevel = null)
    {
        var pmEmail = (string.IsNullOrWhiteSpace(project.ProjectManagerEmail) ? project.ProjectManager : project.ProjectManagerEmail)
            ?.Trim().ToLowerInvariant() ?? string.Empty;

        var parts = new List<string>
        {
            project.ProjectName ?? string.Empty,
            project.CustomerName ?? string.Empty,
            pmEmail,
        };
        parts.AddRange(GranteeEmails(project, leaderEmail, adminEmails));
        // The PIC is also written to its own column, so it must count even when it adds no new grantee (the
        // PIC is often the PM or already on the team). Only when present, so projects without one keep
        // their existing signatures.
        if (!string.IsNullOrWhiteSpace(project.PrincipalInChargeEmail))
        {
            parts.Add("pic:" + project.PrincipalInChargeEmail.Trim().ToLowerInvariant());
        }
        // Only when set, so practices without a Members grant keep their existing signatures (no re-sync).
        if (!string.IsNullOrWhiteSpace(membersPermissionLevel))
        {
            parts.Add("members:" + membersPermissionLevel.Trim().ToLowerInvariant());
        }

        var payload = string.Join(Delimiter, parts);
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(payload));
        return Convert.ToHexString(bytes);
    }
}
