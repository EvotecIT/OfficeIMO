using OfficeIMO.Markdown;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>Controls semantic conversion from Markdown into an OfficeIMO RTF document.</summary>
public sealed class MarkdownRtfExportOptions {
    /// <summary>Markdown reader options used when parsing Markdown text.</summary>
    public MarkdownLoadOptions? ReaderOptions { get; set; }

    /// <summary>Preserves raw HTML as visible text. Otherwise raw HTML is omitted with a diagnostic.</summary>
    public bool PreserveRawHtmlAsText { get; set; }

    /// <summary>Maximum nested Markdown list depth converted into RTF list levels. Default: 32.</summary>
    public int MaxListNestingDepth { get; set; } = 32;

    /// <summary>Maximum number of cells allocated for one converted Markdown table. Default: 100,000.</summary>
    public int MaxTableCells { get; set; } = 100_000;

    /// <summary>RTF writer settings used when the converted document is saved.</summary>
    public OfficeIMO.Rtf.RtfSaveOptions? SaveOptions { get; set; }

    /// <summary>Encoding used when the RTF is saved. Null uses the RTF writer default.</summary>
    public System.Text.Encoding? Encoding { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; nested reader and writer settings are shared.</summary>
    public MarkdownRtfExportOptions Clone() => (MarkdownRtfExportOptions)MemberwiseClone();
}
