using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRouteSettings {
    /// <summary>Native logical Markdown projection settings for PDF input.</summary>
    public PdfLogicalMarkdownOptions? PdfMarkdown { get; set; }

    private static PdfLogicalMarkdownOptions? ClonePdfMarkdownOptions(PdfLogicalMarkdownOptions? options) => options == null
        ? null
        : new PdfLogicalMarkdownOptions {
            IncludePageSeparators = options.IncludePageSeparators,
            IncludeImagePlaceholders = options.IncludeImagePlaceholders,
            IncludeLinkAnnotations = options.IncludeLinkAnnotations,
            IncludeFormWidgets = options.IncludeFormWidgets,
            AlignNumericTableColumns = options.AlignNumericTableColumns,
            PageSeparator = options.PageSeparator
        };
}
