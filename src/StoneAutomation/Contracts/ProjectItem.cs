namespace StoneAutomation.Contracts;

/// <summary>
/// Унифицированная модель изделия/позиции альбома, получаемой из Excel.
/// Эта структура будет использоваться будущим OrderProcessor.
/// </summary>
public sealed record ProjectItem
{
    public required string ModelPath { get; init; }

    public string? Code { get; init; }

    public string? ProjectName { get; init; }

    public string? DrawingName { get; init; }

    public string? OrganizationName { get; init; }

    public string? Sheet { get; init; }

    public string? Sheets { get; init; }

    /// <summary>
    /// Поля штампа и дополнительные prompt-значения в формате "Ключ -> Значение".
    /// Ключи предполагаются case-insensitive.
    /// </summary>
    public IReadOnlyDictionary<string, string> Prompts { get; init; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}
