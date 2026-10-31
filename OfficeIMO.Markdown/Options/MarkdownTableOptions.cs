namespace OfficeIMO.Markdown;

/// <summary>Controls Markdown table projection and optional content-based column alignment.</summary>
public sealed class MarkdownTableOptions : OfficeTableOptions {
    /// <summary>Right-aligns columns whose values are predominantly numeric. Defaults to true.</summary>
    public bool AlignNumeric { get; set; } = true;

    /// <summary>Centers columns whose values are predominantly dates. Defaults to true.</summary>
    public bool AlignDates { get; set; } = true;
}
