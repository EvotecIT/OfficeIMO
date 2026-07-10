namespace OfficeIMO.Reader.Latex;

/// <summary>Options for adapting native LaTeX documents to Reader chunks.</summary>
public sealed class ReaderLatexOptions {
    /// <summary>Emits semantic block chunks instead of whole-document chunks.</summary>
    public bool ChunkByBlock { get; set; } = true;
    /// <summary>Includes parser and conversion diagnostics.</summary>
    public bool IncludeDiagnostics { get; set; } = true;
    /// <summary>Native parse options.</summary>
    public LatexParseOptions ParseOptions { get; set; } = new LatexParseOptions();
    /// <summary>Markdown projection options.</summary>
    public LatexToMarkdownOptions MarkdownOptions { get; set; } = new LatexToMarkdownOptions();
}

internal static class ReaderLatexOptionsCloner {
    internal static ReaderLatexOptions Clone(ReaderLatexOptions? options) {
        ReaderLatexOptions source = options ?? new ReaderLatexOptions();
        LatexParseOptions parse = source.ParseOptions ?? new LatexParseOptions();
        LatexToMarkdownOptions markdown = source.MarkdownOptions ?? new LatexToMarkdownOptions();
        return new ReaderLatexOptions {
            ChunkByBlock = source.ChunkByBlock,
            IncludeDiagnostics = source.IncludeDiagnostics,
            ParseOptions = new LatexParseOptions {
                Profile = parse.Profile,
                MaximumInputLength = parse.MaximumInputLength,
                MaximumTokenCount = parse.MaximumTokenCount,
                MaximumNestingDepth = parse.MaximumNestingDepth,
                MacroExpansion = parse.MacroExpansion,
                MaximumExpansionDepth = parse.MaximumExpansionDepth,
                MaximumExpansionLength = parse.MaximumExpansionLength
            },
            MarkdownOptions = new LatexToMarkdownOptions {
                PreserveUnsupportedAsSource = markdown.PreserveUnsupportedAsSource,
                IncludePreambleAsFrontMatter = markdown.IncludePreambleAsFrontMatter
            }
        };
    }
}
