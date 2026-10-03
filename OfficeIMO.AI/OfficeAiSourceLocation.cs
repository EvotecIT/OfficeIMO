using OfficeIMO.Reader;
using System.Text.Json.Serialization;

namespace OfficeIMO.AI;

/// <summary>Immutable source coordinates supplied by Reader and bound to the evidence snapshot.</summary>
public sealed record OfficeAiSourceLocation {
    /// <summary>File or virtual container path.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Path { get; init; }
    /// <summary>One-based source page, when known.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Page { get; init; }
    /// <summary>One-based presentation slide, when known.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Slide { get; init; }
    /// <summary>Spreadsheet sheet name.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Sheet { get; init; }
    /// <summary>Original source cell or range descriptor, when Reader supplies one.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? A1Range { get; init; }
    /// <summary>Zero-based table index within the closest source container.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? TableIndex { get; init; }
    /// <summary>One-based data-row ordinal within the evidence table; null for its header.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? RowIndex { get; init; }
    /// <summary>Producer-defined source block index.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SourceBlockIndex { get; init; }
    /// <summary>One-based original start line, when known.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? StartLine { get; init; }
    /// <summary>One-based original end line, when known.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? EndLine { get; init; }

    internal static OfficeAiSourceLocation FromReader(ReaderLocation? source, string? path, int? rowIndex) => new() {
        Path = source?.Path ?? path, Page = source?.Page, Slide = source?.Slide, Sheet = source?.Sheet,
        A1Range = source?.A1Range, TableIndex = source?.TableIndex, RowIndex = rowIndex,
        SourceBlockIndex = source?.SourceBlockIndex, StartLine = source?.StartLine, EndLine = source?.EndLine
    };
}
