using OfficeIMO.Markdown;

namespace OfficeIMO.Reader.Markdown;

/// <summary>Controls Markdown parsing and semantic chunking.</summary>
public sealed class ReaderMarkdownOptions {
    /// <summary>Starts a new chunk at headings when possible.</summary>
    public bool ChunkByHeadings { get; set; } = true;

    /// <summary>
    /// Markdown parser configuration. Registration snapshots this object with
    /// <see cref="MarkdownLoadOptions.Clone"/>.
    /// </summary>
    public MarkdownLoadOptions? ParserOptions { get; set; }
}
