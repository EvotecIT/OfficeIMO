namespace OfficeIMO.Reader.AsciiDoc;

/// <summary>Options for adapting native AsciiDoc documents to Reader chunks.</summary>
public sealed class ReaderAsciiDocOptions {
    /// <summary>Emits one logical chunk per supported source block. Defaults to true.</summary>
    public bool ChunkByBlock { get; set; } = true;

    /// <summary>Includes source comments as chunks. Defaults to false.</summary>
    public bool IncludeComments { get; set; }

    /// <summary>Includes document attribute entries as chunks. Defaults to false.</summary>
    public bool IncludeAttributes { get; set; }

    /// <summary>Includes parser and conversion diagnostics as chunk warnings. Defaults to true.</summary>
    public bool IncludeDiagnostics { get; set; } = true;

    /// <summary>Native parser options.</summary>
    public AsciiDocParseOptions ParseOptions { get; set; } = new AsciiDocParseOptions();

    /// <summary>Markdown projection options.</summary>
    public AsciiDocToMarkdownOptions MarkdownOptions { get; set; } = new AsciiDocToMarkdownOptions();
}

internal static class ReaderAsciiDocOptionsCloner {
    internal static ReaderAsciiDocOptions Clone(ReaderAsciiDocOptions? options) {
        ReaderAsciiDocOptions source = options ?? new ReaderAsciiDocOptions();
        AsciiDocParseOptions parse = source.ParseOptions ?? new AsciiDocParseOptions();
        AsciiDocToMarkdownOptions markdown = source.MarkdownOptions ?? new AsciiDocToMarkdownOptions();
        return new ReaderAsciiDocOptions {
            ChunkByBlock = source.ChunkByBlock,
            IncludeComments = source.IncludeComments,
            IncludeAttributes = source.IncludeAttributes,
            IncludeDiagnostics = source.IncludeDiagnostics,
            ParseOptions = new AsciiDocParseOptions {
                MaximumInputLength = parse.MaximumInputLength,
                MaximumBlockCount = parse.MaximumBlockCount,
                MaximumInlineNestingDepth = parse.MaximumInlineNestingDepth,
                MaximumInlineNodeCount = parse.MaximumInlineNodeCount
            },
            MarkdownOptions = new AsciiDocToMarkdownOptions {
                IncludeDocumentAttributesAsFrontMatter = markdown.IncludeDocumentAttributesAsFrontMatter,
                PreserveUnsupportedAsSource = markdown.PreserveUnsupportedAsSource,
                PreserveCommentsAsSource = markdown.PreserveCommentsAsSource,
                ExpandDocumentAttributes = markdown.ExpandDocumentAttributes,
                UndefinedAttributeBehavior = markdown.UndefinedAttributeBehavior
            }
        };
    }
}
