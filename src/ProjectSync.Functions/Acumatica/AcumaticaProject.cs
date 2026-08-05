namespace ProjectSync.Acumatica;

/// <summary>A single project row read from the Acumatica Generic Inquiry.</summary>
public sealed record AcumaticaProject
{
    public required string ProjectId { get; init; }
    public string? ProjectName { get; init; }
    public string? CustomerName { get; init; }
    public string? ProjectManager { get; init; }
    public string? Practice { get; init; }
    public DateTimeOffset? CreatedDateTime { get; init; }
}
