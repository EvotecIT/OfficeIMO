namespace OfficeIMO.Markdown;

/// <summary>
/// Source-backed token or payload field owned by a native markdown block.
/// </summary>
public sealed class MarkdownNativeBlockSourceField {
    internal MarkdownNativeBlockSourceField(
        string name,
        string? value,
        MarkdownSourceSpan sourceSpan,
        MarkdownNativeBlock block,
        int index = -1) {
        Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentException("Field name is required.", nameof(name)) : name;
        Value = value;
        SourceSpan = sourceSpan;
        Block = block ?? throw new ArgumentNullException(nameof(block));
        Index = index;
    }

    /// <summary>Stable field name such as <c>level</c>, <c>infoString</c>, or <c>calloutKind</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field in the normalized markdown source.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>Native block that owns this source field.</summary>
    public MarkdownNativeBlock Block { get; }

    /// <summary>Zero-based occurrence index for repeated fields, or <c>-1</c> for singular fields.</summary>
    public int Index { get; }
}
