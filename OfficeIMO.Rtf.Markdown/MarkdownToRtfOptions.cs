using OfficeIMO.Markdown;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>Controls semantic conversion from Markdown into an OfficeIMO RTF document.</summary>
public sealed class MarkdownToRtfOptions {
    /// <summary>Markdown reader options used when parsing Markdown text.</summary>
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Preserves raw HTML as visible text. Otherwise raw HTML is omitted with a diagnostic.</summary>
    public bool PreserveRawHtmlAsText { get; set; }
}
