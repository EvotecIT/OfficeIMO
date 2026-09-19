namespace OfficeIMO.Pdf;

/// <summary>
/// Describes page content that is and is not in scope when logical PDF tables are extracted.
/// </summary>
public sealed class PdfTableExtractionScopeReport {
    internal PdfTableExtractionScopeReport(
        int sourcePageCount,
        int pagesWithTables,
        int detectedTableCount,
        int nonTableTextBlockCount,
        int vectorPrimitiveCount,
        int imageCount,
        int linkCount,
        int formWidgetCount,
        int formFieldCount,
        bool hasAcroFormXfa,
        int annotationCount,
        int pageActionCount,
        int catalogActionCount,
        bool hasOpenAction,
        int documentActionCount,
        int optionalContentGroupCount,
        int pagesWithOptionalContent,
        int interactiveMediaAnnotationCount,
        int unplacedFormFieldCount,
        int outlineCount,
        int attachmentCount,
        bool hasTaggedContent,
        bool analysisTruncated,
        bool hasSourceSecurityState = false,
        int pageLabelCount = 0,
        bool hasDocumentMetadata = false) {
        SourcePageCount = sourcePageCount;
        PagesWithTables = pagesWithTables;
        DetectedTableCount = detectedTableCount;
        NonTableTextBlockCount = nonTableTextBlockCount;
        VectorPrimitiveCount = vectorPrimitiveCount;
        ImageCount = imageCount;
        LinkCount = linkCount;
        FormWidgetCount = formWidgetCount;
        FormFieldCount = formFieldCount;
        HasAcroFormXfa = hasAcroFormXfa;
        AnnotationCount = annotationCount;
        PageActionCount = pageActionCount;
        CatalogActionCount = catalogActionCount;
        HasOpenAction = hasOpenAction;
        DocumentActionCount = documentActionCount;
        OptionalContentGroupCount = optionalContentGroupCount;
        PagesWithOptionalContent = pagesWithOptionalContent;
        InteractiveMediaAnnotationCount = interactiveMediaAnnotationCount;
        UnplacedFormFieldCount = unplacedFormFieldCount;
        OutlineCount = outlineCount;
        AttachmentCount = attachmentCount;
        HasTaggedContent = hasTaggedContent;
        AnalysisTruncated = analysisTruncated;
        HasSourceSecurityState = hasSourceSecurityState;
        PageLabelCount = pageLabelCount;
        HasDocumentMetadata = hasDocumentMetadata;
    }

    /// <summary>Number of logical source pages inspected.</summary>
    public int SourcePageCount { get; }

    /// <summary>Number of source pages on which at least one logical table was detected.</summary>
    public int PagesWithTables { get; }

    /// <summary>Number of source pages on which no logical table was detected.</summary>
    public int PagesWithoutTables => SourcePageCount - PagesWithTables;

    /// <summary>Total number of logical tables detected on the inspected pages.</summary>
    public int DetectedTableCount { get; }

    /// <summary>Number of visible text blocks that were not represented by a detected table.</summary>
    public int NonTableTextBlockCount { get; }

    /// <summary>
    /// Number of source vector drawing primitives. Table-only adapters do not import the original vector artwork,
    /// even when its geometry contributed to logical table detection.
    /// </summary>
    public int VectorPrimitiveCount { get; }

    /// <summary>Number of source images with at least one visible page placement, which table-only adapters do not import.</summary>
    public int ImageCount { get; }

    /// <summary>Number of source link annotations, which table-only adapters do not import.</summary>
    public int LinkCount { get; }

    /// <summary>Number of source form widgets, which table-only adapters do not import.</summary>
    public int FormWidgetCount { get; }

    /// <summary>Number of document-level AcroForm fields, which table-only adapters do not import.</summary>
    public int FormFieldCount { get; }

    /// <summary>Form fields with no widgets or with a widget not attached to a readable page.</summary>
    public int UnplacedFormFieldCount { get; }

    /// <summary>Whether the document has an XFA form definition, which table-only adapters do not import.</summary>
    public bool HasAcroFormXfa { get; }

    /// <summary>
    /// Number of distinct form signals outside table-only output. Fields and widgets can describe the same
    /// controls, so the larger count is used, with an XFA definition counted separately.
    /// </summary>
    public int FormContentCount => Math.Max(FormWidgetCount, FormFieldCount) + (HasAcroFormXfa ? 1 : 0);

    /// <summary>
    /// Number of generic source annotation records, which table-only adapters do not import.
    /// Link and widget annotations may also appear in their dedicated counts, so these counts are not additive.
    /// </summary>
    public int AnnotationCount { get; }

    /// <summary>Number of source page actions, which table-only adapters do not import.</summary>
    public int PageActionCount { get; }

    /// <summary>Number of source catalog actions, which table-only adapters do not import.</summary>
    public int CatalogActionCount { get; }

    /// <summary>Whether the source has a readable document-open destination or GoTo action.</summary>
    public bool HasOpenAction { get; }

    /// <summary>Total distinct catalog and readable document-open actions outside table-only output.</summary>
    public int DocumentActionCount { get; }

    /// <summary>Number of optional-content groups, which table-only adapters do not preserve as editable groups or layers.</summary>
    public int OptionalContentGroupCount { get; }

    /// <summary>Number of inspected source pages that actually use optional content.</summary>
    public int PagesWithOptionalContent { get; }

    /// <summary>Number of movie, sound, screen, rich-media, or 3D annotations, which table-only adapters do not import as animations or media.</summary>
    public int InteractiveMediaAnnotationCount { get; }

    /// <summary>Document outline entries outside table-only output.</summary>
    public int OutlineCount { get; }

    /// <summary>Embedded document attachments outside table-only output.</summary>
    public int AttachmentCount { get; }

    /// <summary>Whether the source has a tagged PDF structure tree that table-only output does not preserve.</summary>
    public bool HasTaggedContent { get; }

    /// <summary>Whether source encryption, signatures, permissions, or revision state cannot be carried into table output.</summary>
    public bool HasSourceSecurityState { get; }

    /// <summary>Number of selected-source page-label rules not carried into table output.</summary>
    public int PageLabelCount { get; }

    /// <summary>Whether populated source Info or XMP document metadata is outside table-only output.</summary>
    public bool HasDocumentMetadata { get; }

    /// <summary>
    /// True when bounded text/table correlation stopped before every visible text block could be classified.
    /// Exact omission counts describe only blocks classified before the limit was reached.
    /// </summary>
    public bool AnalysisTruncated { get; }

    /// <summary>
    /// Gets whether visible or interactive page or document content existed outside the detected tables.
    /// This is expected for table-only extraction and is separate from truncation within a table.
    /// </summary>
    public bool HasOmittedPageContent =>
        AnalysisTruncated ||
        NonTableTextBlockCount > 0 ||
        VectorPrimitiveCount > 0 ||
        ImageCount > 0 ||
        LinkCount > 0 ||
        FormContentCount > 0 ||
        AnnotationCount > 0 ||
        PageActionCount > 0 ||
        DocumentActionCount > 0 ||
        OutlineCount > 0 ||
        AttachmentCount > 0 ||
        HasTaggedContent ||
        HasSourceSecurityState ||
        PageLabelCount > 0 ||
        HasDocumentMetadata ||
        PagesWithOptionalContent > 0 ||
        InteractiveMediaAnnotationCount > 0;
}
