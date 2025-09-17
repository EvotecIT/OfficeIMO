namespace OfficeIMO.Markdown;

/// <summary>
/// Contract for a single block parser used by the Markdown reader pipeline.
/// </summary>
public interface IMarkdownBlockParser {
    /// <summary>
    /// Attempts to parse a block starting at the specified line.
    /// </summary>
    /// <param name="lines">All lines of the document.</param>
    /// <param name="i">Current index within <paramref name="lines"/>. Implementations advance it when they consume lines.</param>
    /// <param name="options">Reader options controlling parsing behavior.</param>
    /// <param name="doc">Markdown document model being built.</param>
    /// <param name="state">Mutable state shared across parsers.</param>
    /// <returns>true if a block was parsed; otherwise false.</returns>
    bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state);
}
