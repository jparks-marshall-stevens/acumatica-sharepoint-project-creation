namespace ProjectSync.Acumatica;

/// <summary>A single project row read from the Acumatica Generic Inquiry.</summary>
public sealed record AcumaticaProject
{
    public required string ProjectId { get; init; }
    public string? ProjectName { get; init; }
    public string? CustomerName { get; init; }
    public string? ProjectManager { get; init; }

    /// <summary>Project manager's email/UPN, used to resolve a SharePoint People field. Optional.</summary>
    public string? ProjectManagerEmail { get; init; }

    public string? Practice { get; init; }
    public DateTimeOffset? CreatedDateTime { get; init; }

    /// <summary>Emails of the project's team/employees (from the team GI). Granted Edit alongside the PM + leader.</summary>
    public IReadOnlyList<string> TeamEmails { get; init; } = Array.Empty<string>();
}
