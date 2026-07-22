using OfficeIMO.Markdown;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>Controls semantic conversion from Markdown into an OfficeIMO RTF document.</summary>
public sealed class MarkdownToRtfOptions {
    /// <summary>Markdown reader options used when parsing Markdown text.</summary>
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Preserves raw HTML as visible text. Otherwise raw HTML is omitted with a diagnostic.</summary>
    public bool PreserveRawHtmlAsText { get; set; }

    /// <summary>Maximum nested Markdown list depth converted into RTF list levels. Default: 32.</summary>
    public int MaxListNestingDepth { get; set; } = 32;

    /// <summary>Maximum number of cells allocated for one converted Markdown table. Default: 100,000.</summary>
    public int MaxTableCells { get; set; } = 100_000;
}
