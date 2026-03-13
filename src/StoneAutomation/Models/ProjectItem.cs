namespace StoneDrawingAuto.Core.Models;

public sealed record ProjectItem
{
    public required string ModelPath { get; init; }
    public string? Code { get; init; }
    public string? ProjectName { get; init; }
    public string? DrawingName { get; init; }
    public IReadOnlyDictionary<string, string> Prompts { get; init; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}
