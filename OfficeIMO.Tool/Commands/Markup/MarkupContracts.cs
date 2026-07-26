using System.Text.Json.Serialization;

namespace OfficeIMO.Tool.Commands.Markup;

internal sealed record MarkupEnvelope(
    OfficeMarkupDocumentDto Document,
    IReadOnlyList<OfficeMarkupDiagnosticDto> Diagnostics);

internal sealed record ValidationEnvelope(
    IReadOnlyList<OfficeMarkupDiagnosticDto> Diagnostics,
    bool HasErrors);

internal sealed record ExportEnvelope(string OutputPath, string Target);

internal sealed class OfficeMarkupDocumentDto {
    public string Profile { get; set; } = string.Empty;
    public Dictionary<string, string> Metadata { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<OfficeMarkupBlockDto> Blocks { get; set; } = [];
}

internal sealed class OfficeMarkupBlockDto {
    public string Kind { get; set; } = string.Empty;
    public Dictionary<string, string> Attributes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? SourceText { get; set; }
    public string? Text { get; set; }
    public int? Level { get; set; }
    public bool? Ordered { get; set; }
    public int? Start { get; set; }
    public List<OfficeMarkupListItemDto>? Items { get; set; }
    public string? Language { get; set; }
    public string? Content { get; set; }
    public string? Source { get; set; }
    public string? Alt { get; set; }
    public string? Title { get; set; }
    public double? Width { get; set; }
    public double? Height { get; set; }
    public List<string>? Headers { get; set; }
    public List<List<string>>? Rows { get; set; }
    public bool? RenderAsImage { get; set; }
    public string? Layout { get; set; }
    public string? Section { get; set; }
    public string? Transition { get; set; }
    public string? Background { get; set; }
    public string? Notes { get; set; }
    public string? Placement { get; set; }
    public int? Columns { get; set; }
    public List<OfficeMarkupBlockDto>? Blocks { get; set; }
    public string? Name { get; set; }
    public string? PageSize { get; set; }
    public string? Orientation { get; set; }
    public int? MinLevel { get; set; }
    public int? MaxLevel { get; set; }
    public string? Address { get; set; }
    public string? Sheet { get; set; }
    public string? Cell { get; set; }
    public string? Expression { get; set; }
    public string? Range { get; set; }
    public bool? HasHeader { get; set; }
    public string? ChartType { get; set; }
    public string? Target { get; set; }
    public string? Style { get; set; }
    public string? NumberFormat { get; set; }
    public string? Gap { get; set; }
    public string? ColumnKind { get; set; }
    public string? WidthText { get; set; }
    public OfficeMarkupPlacementDto? Position { get; set; }
    public OfficeMarkupResolvedStyleDto? ResolvedStyle { get; set; }
    public OfficeMarkupTransitionDto? TransitionDetails { get; set; }
    public string? Command { get; set; }
    public string? Body { get; set; }
    public string? Markdown { get; set; }
}

internal sealed class OfficeMarkupTransitionDto {
    public string? RawText { get; set; }
    public string? Effect { get; set; }
    public string? ResolvedIdentifier { get; set; }
    public Dictionary<string, string> Attributes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

internal sealed class OfficeMarkupResolvedStyleDto {
    public string? Name { get; set; }
    public string? FontName { get; set; }
    public int? FontSize { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? TextColor { get; set; }
    public string? FillColor { get; set; }
    public string? BorderColor { get; set; }
    public string? TextAlign { get; set; }
}

internal sealed class OfficeMarkupPlacementDto {
    public string? X { get; set; }
    public string? Y { get; set; }
    public string? Width { get; set; }
    public string? Height { get; set; }
}

internal sealed class OfficeMarkupListItemDto {
    public string Text { get; set; } = string.Empty;
    public bool IsTask { get; set; }
    public bool IsChecked { get; set; }
    public List<OfficeMarkupBlockDto> Blocks { get; set; } = [];
}

internal sealed class OfficeMarkupDiagnosticDto {
    public string Severity { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? NodeKind { get; set; }
    public string? NodeSourceText { get; set; }
}

[JsonSourceGenerationOptions(
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(MarkupEnvelope))]
[JsonSerializable(typeof(ValidationEnvelope))]
[JsonSerializable(typeof(ExportEnvelope))]
internal sealed partial class MarkupJsonSerializerContext : JsonSerializerContext;
