namespace OfficeIMO.Markdown;

/// <summary>
/// Contract for a single block parser used by the Markdown reader pipeline.
/// Implementations examine <paramref name="lines"/> at index <paramref name="i"/> and, when matched,
/// append one or more blocks to <paramref name="doc"/> and advance <paramref name="i"/> accordingly.
/// </summary>
public interface IMarkdownBlockParser {
    bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state);
}
