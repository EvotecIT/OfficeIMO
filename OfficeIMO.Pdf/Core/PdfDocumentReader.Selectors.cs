namespace OfficeIMO.Pdf;

internal sealed partial class PdfDocumentReader {
    /// <summary>Extracts plain text from pages resolved by a document-relative selector.</summary>
    public string Text(OfficePageSelector selector, PdfTextLayoutOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Text(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to extract plain text from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<string> TextResult(OfficePageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts text per page in the order resolved by a document-relative selector.</summary>
    public IReadOnlyList<string> TextByPage(OfficePageSelector selector, PdfLoadOptions? readOptions = null) =>
        TextByPage(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract text per page resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<string>> TextByPageResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts Markdown from pages resolved by a document-relative selector.</summary>
    public string Markdown(OfficePageSelector selector, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfLoadOptions? readOptions = null) =>
        Markdown(ResolveSelector(selector, readOptions), options, markdownOptions, readOptions);

    /// <summary>Attempts to extract Markdown from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<string> MarkdownResult(OfficePageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(selector, layoutOptions, markdownOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Builds the logical model for pages resolved by a document-relative selector.</summary>
    public PdfDocumentReadResult Logical(OfficePageSelector selector, PdfTextLayoutOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Logical(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to build the logical model for pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocumentReadResult> LogicalResult(OfficePageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts logical text blocks from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks(OfficePageSelector selector, PdfTextLayoutOptions? options = null, PdfLoadOptions? readOptions = null) =>
        TextBlocks(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to extract logical text blocks from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> TextBlocksResult(OfficePageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract logical text blocks", PdfPreflightCapability.ReadLogicalObjects, () => TextBlocks(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts images from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfExtractedImage> Images(OfficePageSelector selector, PdfLoadOptions? readOptions = null) =>
        Images(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract images from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<PdfExtractedImage>> ImagesResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, () => Images(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts image placements from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements(OfficePageSelector selector, PdfLoadOptions? readOptions = null) =>
        ImagePlacements(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract image placements from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<PdfImagePlacement>> ImagePlacementsResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract image placements", PdfPreflightCapability.ExtractImages, () => ImagePlacements(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Inspects unique fonts declared by pages resolved by a document-relative selector.</summary>
    public PdfFontInventory Fonts(
        OfficePageSelector selector,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) =>
        Fonts(ResolveSelector(selector, readOptions), inspectionOptions, readOptions);

    /// <summary>Attempts to inspect fonts on pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<PdfFontInventory> FontsResult(
        OfficePageSelector selector,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Inspect fonts",
            PdfPreflightCapability.ReadLogicalObjects,
            () => Fonts(selector, inspectionOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Recovers conservative paragraph continuation groups from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalParagraphContinuationGroup> ParagraphContinuations(
        OfficePageSelector selector,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) =>
        ParagraphContinuations(ResolveSelector(selector, readOptions), continuationOptions, layoutOptions, readOptions);

    /// <summary>Attempts to recover paragraph continuation groups from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> ParagraphContinuationsResult(
        OfficePageSelector selector,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Recover paragraph continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => ParagraphContinuations(selector, continuationOptions, layoutOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Recovers bounded table continuation groups from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> TableContinuations(
        OfficePageSelector selector,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) =>
        TableContinuations(ResolveSelector(selector, readOptions), continuationOptions, layoutOptions, readOptions);

    /// <summary>Attempts to recover table continuation groups from pages resolved by a document-relative selector.</summary>
    internal PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> TableContinuationsResult(
        OfficePageSelector selector,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Recover table continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => TableContinuations(selector, continuationOptions, layoutOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        OfficePageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(selector, nameof(selector));
        return PdfPageImageRenderer.RenderPages(
            token => _document.GetReadDocument(ResolveReadOptions(readOptions), token),
            selector,
            options,
            cancellationToken);
    }

    private PdfPageSelection ResolveSelector(OfficePageSelector selector, PdfLoadOptions? readOptions) {
        Guard.NotNull(selector, nameof(selector));
        int pageCount = _document.Inspect(ResolveReadOptions(readOptions)).PageCount;
        if (pageCount < 1) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        return selector.ResolveSelection(pageCount);
    }
}
