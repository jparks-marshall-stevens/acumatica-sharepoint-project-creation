using ProjectSync.Acumatica;
using ProjectSync.SharePoint;
using Xunit;

namespace ProjectSync.Functions.Tests;

public class ReconcileSignatureTests
{
    private static AcumaticaProject Project() => new()
    {
        ProjectId = "P1",
        ProjectName = "Name",
        CustomerName = "Customer",
        ProjectManagerEmail = "pm@x.com",
        TeamEmails = new[] { "team@x.com" },
    };

    [Fact]
    public void GranteeEmails_IncludesPracticeAdmins()
    {
        var grantees = ReconcileSignature.GranteeEmails(
            Project(), "leader@x.com", new[] { "admin@x.com" });

        Assert.Contains("admin@x.com", grantees);
        Assert.Contains("leader@x.com", grantees);
        Assert.Contains("pm@x.com", grantees);
        Assert.Contains("team@x.com", grantees);
    }

    [Fact]
    public void Compute_ChangesWhenAnAdminIsAdded()
    {
        // Adding a practice admin must change the signature, so the daily reconcile re-applies
        // permissions (grants the admin) instead of treating the set as unchanged.
        var without = ReconcileSignature.Compute(Project(), "leader@x.com");
        var with = ReconcileSignature.Compute(Project(), "leader@x.com", new[] { "admin@x.com" });

        Assert.NotEqual(without, with);
    }

    [Fact]
    public void Compute_IsStableRegardlessOfAdminOrderOrCase()
    {
        var a = ReconcileSignature.Compute(Project(), "leader@x.com", new[] { "Admin@X.com", "b@x.com" });
        var b = ReconcileSignature.Compute(Project(), "leader@x.com", new[] { "b@x.com", "admin@x.com" });

        Assert.Equal(a, b);
    }

    [Fact]
    public void NoMembersLevel_KeepsTheExistingSignature() =>
        Assert.Equal(ReconcileSignature.Compute(Project(), "leader@x.com"),
            ReconcileSignature.Compute(Project(), "leader@x.com", null, membersPermissionLevel: " "));

    [Fact]
    public void MembersLevel_ChangesTheSignature_SoExistingRoomsResync()
    {
        var baseline = ReconcileSignature.Compute(Project(), "leader@x.com");
        var edit = ReconcileSignature.Compute(Project(), "leader@x.com", null, "Edit");
        Assert.NotEqual(baseline, edit);
        Assert.Equal(edit, ReconcileSignature.Compute(Project(), "leader@x.com", null, " edit "));
        Assert.NotEqual(edit, ReconcileSignature.Compute(Project(), "leader@x.com", null, "Read"));
    }
}
