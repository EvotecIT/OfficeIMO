namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Options for converting native AsciiDoc blocks to Markdown.</summary>
public sealed class AsciiDocToMarkdownOptions {
    /// <summary>Maps set document attributes to YAML front matter. Defaults to true.</summary>
    public bool IncludeDocumentAttributesAsFrontMatter { get; set; } = true;

    /// <summary>Preserves unsupported semantic blocks as fenced <c>asciidoc</c> source. Defaults to true.</summary>
    public bool PreserveUnsupportedAsSource { get; set; } = true;

    /// <summary>Preserves comments as fenced <c>asciidoc</c> source instead of omitting them.</summary>
    public bool PreserveCommentsAsSource { get; set; }

    /// <summary>Expands document attribute references in converted inline content. Defaults to true.</summary>
    public bool ExpandDocumentAttributes { get; set; } = true;

    /// <summary>Behavior for undefined attributes when expansion is enabled.</summary>
    public AsciiDocUndefinedAttributeBehavior UndefinedAttributeBehavior { get; set; } = AsciiDocUndefinedAttributeBehavior.Preserve;

    /// <summary>Maximum recursively converted compound-block depth. Defaults to 64.</summary>
    public int MaximumBlockNestingDepth { get; set; } = 64;

    /// <summary>Optional document reference snapshot when converting an individual block.</summary>
    public AsciiDocReferenceCatalog? References { get; set; }
    internal HashSet<string> UsedFootnotes { get; } = new HashSet<string>(StringComparer.Ordinal);

    internal AsciiDocToMarkdownOptions WithReferences(AsciiDocReferenceCatalog references) => new AsciiDocToMarkdownOptions {
        IncludeDocumentAttributesAsFrontMatter = IncludeDocumentAttributesAsFrontMatter,
        PreserveUnsupportedAsSource = PreserveUnsupportedAsSource,
        PreserveCommentsAsSource = PreserveCommentsAsSource,
        ExpandDocumentAttributes = ExpandDocumentAttributes,
        UndefinedAttributeBehavior = UndefinedAttributeBehavior,
        MaximumBlockNestingDepth = MaximumBlockNestingDepth,
        References = References ?? references
    };
}
