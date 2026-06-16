using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>
/// Options for importing parser-supported PDF content into the OfficeIMO RTF document model.
/// </summary>
/// <remarks>
/// PDF import is semantic text extraction over the first-party logical PDF reader. It preserves
/// supported document metadata, page breaks, headings, grouped paragraphs, and list markers, but
/// it is not a visual reconstruction of arbitrary fixed-layout PDF content.
/// </remarks>
public sealed class PdfRtfReadOptions {
    /// <summary>Logical PDF layout options used while grouping positioned PDF text into semantic objects.</summary>
    public PdfCore.PdfTextLayoutOptions? LayoutOptions { get; set; }

    /// <summary>Whether PDF Info dictionary metadata should be copied into the RTF info destination.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>Whether page transitions should be represented by RTF page-break-before paragraphs.</summary>
    public bool PreservePageBreaks { get; set; } = true;

    /// <summary>Whether empty PDF pages should produce an empty RTF paragraph when page breaks are preserved.</summary>
    public bool IncludeEmptyPages { get; set; }

    /// <summary>Whether logical heading lines should be imported as RTF heading-like paragraphs.</summary>
    public bool ImportHeadings { get; set; } = true;

    /// <summary>Whether logical list items should be imported as RTF list paragraphs with fallback marker text.</summary>
    public bool ImportLists { get; set; } = true;

    /// <summary>Whether common Heading 1-3 stylesheet entries should be created for imported headings.</summary>
    public bool CreateHeadingStyles { get; set; } = true;

    /// <summary>Creates a reusable copy of this option set.</summary>
    public PdfRtfReadOptions Clone() => new PdfRtfReadOptions {
        LayoutOptions = CloneLayoutOptions(LayoutOptions),
        IncludeMetadata = IncludeMetadata,
        PreservePageBreaks = PreservePageBreaks,
        IncludeEmptyPages = IncludeEmptyPages,
        ImportHeadings = ImportHeadings,
        ImportLists = ImportLists,
        CreateHeadingStyles = CreateHeadingStyles
    };

    private static PdfCore.PdfTextLayoutOptions? CloneLayoutOptions(PdfCore.PdfTextLayoutOptions? options) {
        if (options is null) {
            return null;
        }

        return new PdfCore.PdfTextLayoutOptions {
            MarginLeft = options.MarginLeft,
            MarginRight = options.MarginRight,
            BinWidth = options.BinWidth,
            MinGutterWidth = options.MinGutterWidth,
            LineMergeToleranceEm = options.LineMergeToleranceEm,
            LineMergeMaxPoints = options.LineMergeMaxPoints,
            ForceSingleColumn = options.ForceSingleColumn,
            JoinHyphenationAcrossLines = options.JoinHyphenationAcrossLines,
            IgnoreHeaderHeight = options.IgnoreHeaderHeight,
            IgnoreFooterHeight = options.IgnoreFooterHeight,
            GapSpaceThresholdEm = options.GapSpaceThresholdEm,
            GapGlyphFactor = options.GapGlyphFactor
        };
    }
}
