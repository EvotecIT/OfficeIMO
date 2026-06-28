namespace OfficeIMO.Markdown;

/// <summary>
/// Source-backed token or payload field owned by a reference-style link definition.
/// </summary>
public sealed class MarkdownNativeReferenceLinkDefinitionField {
    internal MarkdownNativeReferenceLinkDefinitionField(
        string name,
        string? value,
        MarkdownSourceSpan sourceSpan,
        MarkdownReferenceLinkDefinition definition) {
        Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentException("Field name is required.", nameof(name)) : name;
        Value = value;
        SourceSpan = sourceSpan;
        Definition = definition ?? throw new ArgumentNullException(nameof(definition));
    }

    /// <summary>Stable field name such as <c>openingMarker</c>, <c>label</c>, <c>separatorMarker</c>, <c>url</c>, or <c>title</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field in the normalized markdown source.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>Reference-style link definition that owns this source field.</summary>
    public MarkdownReferenceLinkDefinition Definition { get; }
}
