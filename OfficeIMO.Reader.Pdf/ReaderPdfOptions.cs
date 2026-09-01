using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// Options for PDF ingestion through the OfficeIMO.Reader adapter.
/// </summary>
public sealed class ReaderPdfOptions {
    /// <summary>
    /// Creates the default PDF reader profile with page-level chunks and wrapper-friendly Markdown.
    /// </summary>
    public static ReaderPdfOptions CreateOfficeIMOProfile() => new ReaderPdfOptions {
        MarkdownOptions = new PdfLogicalMarkdownOptions {
            IncludePageSeparators = false,
            IncludeImagePlaceholders = true,
            IncludeLinkAnnotations = true,
            IncludeFormWidgets = true
        }
    };

    /// <summary>
    /// Optional PDF text layout options used by the logical read model.
    /// </summary>
    public PdfTextLayoutOptions? LayoutOptions { get; set; }

    /// <summary>
    /// Optional inclusive one-based source page ranges. Null reads the full document.
    /// </summary>
    public IReadOnlyList<PdfPageRange>? PageRanges { get; set; }

    /// <summary>
    /// Maximum selected pages processed by the canonical semantic pipeline.
    /// Increase this explicitly when ingesting larger PDF documents.
    /// </summary>
    public int MaxPages { get; set; } = PdfUnderstandingPipelineOptions.DefaultMaxPages;

    /// <summary>
    /// Markdown rendering options used for page chunk content.
    /// </summary>
    public PdfLogicalMarkdownOptions? MarkdownOptions { get; set; }

    /// <summary>
    /// When true, projects conservative cross-page paragraph continuation evidence into document metadata.
    /// Page-local text and chunks remain unchanged. Default: true.
    /// </summary>
    public bool IncludeParagraphContinuationMetadata { get; set; } = true;

    /// <summary>Optional confidence and hyphenation policy for paragraph continuation metadata.</summary>
    public PdfLogicalParagraphContinuationOptions? ParagraphContinuationOptions { get; set; }

    /// <summary>
    /// When true, emits one or more chunks per logical source page. Default: true.
    /// </summary>
    public bool ChunkByPage { get; set; } = true;

    /// <summary>
    /// Creates a defensive copy for handler registration reuse.
    /// </summary>
    public ReaderPdfOptions Clone() => new ReaderPdfOptions {
        LayoutOptions = CloneLayoutOptions(LayoutOptions),
        PageRanges = PageRanges?.ToArray(),
        MaxPages = MaxPages,
        MarkdownOptions = CloneMarkdownOptions(MarkdownOptions),
        IncludeParagraphContinuationMetadata = IncludeParagraphContinuationMetadata,
        ParagraphContinuationOptions = CloneParagraphContinuationOptions(ParagraphContinuationOptions),
        ChunkByPage = ChunkByPage
    };

    internal static PdfTextLayoutOptions? CloneLayoutOptions(PdfTextLayoutOptions? options) {
        if (options is null) return null;

        return new PdfTextLayoutOptions {
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

    internal static PdfLogicalParagraphContinuationOptions? CloneParagraphContinuationOptions(PdfLogicalParagraphContinuationOptions? options) {
        if (options is null) return null;

        return new PdfLogicalParagraphContinuationOptions {
            MergePageContinuations = options.MergePageContinuations,
            RejoinLineEndingHyphens = options.RejoinLineEndingHyphens,
            MaximumSegmentsPerParagraph = options.MaximumSegmentsPerParagraph,
            GeometryTolerancePoints = options.GeometryTolerancePoints,
            MinimumConfidence = options.MinimumConfidence
        };
    }

    internal static PdfLogicalMarkdownOptions? CloneMarkdownOptions(PdfLogicalMarkdownOptions? options) {
        if (options is null) return null;

        return new PdfLogicalMarkdownOptions {
            IncludePageSeparators = options.IncludePageSeparators,
            IncludeImagePlaceholders = options.IncludeImagePlaceholders,
            IncludeLinkAnnotations = options.IncludeLinkAnnotations,
            IncludeFormWidgets = options.IncludeFormWidgets,
            AlignNumericTableColumns = options.AlignNumericTableColumns,
            PageSeparator = options.PageSeparator
        };
    }
}
