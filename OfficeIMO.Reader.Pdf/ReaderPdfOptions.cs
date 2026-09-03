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
    /// Canonical semantic read settings. Null uses <see cref="PdfReadOptions.Default"/>.
    /// When the adapter receives an already reconstructed PDF result, only its page selection is applied.
    /// </summary>
    public PdfReadOptions? ReadOptions { get; set; }

    /// <summary>
    /// Markdown rendering options used for page chunk content.
    /// </summary>
    public PdfLogicalMarkdownOptions? MarkdownOptions { get; set; }

    /// <summary>
    /// When true, projects conservative cross-page paragraph continuation evidence into document metadata.
    /// Page-local text and chunks remain unchanged. Default: true.
    /// </summary>
    public bool IncludeParagraphContinuationMetadata { get; set; } = true;

    /// <summary>Optional confidence and soft-hyphen policy for paragraph continuation metadata.</summary>
    public PdfLogicalParagraphContinuationOptions? ParagraphContinuationOptions { get; set; }

    /// <summary>
    /// When true, emits one or more chunks per logical source page. Default: true.
    /// </summary>
    public bool ChunkByPage { get; set; } = true;

    /// <summary>
    /// Creates a defensive copy for handler registration reuse.
    /// </summary>
    public ReaderPdfOptions Clone() => new ReaderPdfOptions {
        ReadOptions = ReadOptions?.Clone(),
        MarkdownOptions = CloneMarkdownOptions(MarkdownOptions),
        IncludeParagraphContinuationMetadata = IncludeParagraphContinuationMetadata,
        ParagraphContinuationOptions = CloneParagraphContinuationOptions(ParagraphContinuationOptions),
        ChunkByPage = ChunkByPage
    };

    internal static PdfLogicalParagraphContinuationOptions? CloneParagraphContinuationOptions(PdfLogicalParagraphContinuationOptions? options) {
        if (options is null) return null;

        return new PdfLogicalParagraphContinuationOptions {
            MergePageContinuations = options.MergePageContinuations,
            RejoinSoftHyphens = options.RejoinSoftHyphens,
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
