using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Extracts plain text from pages resolved by a document-relative selector.</summary>
    public string Text(PdfPageSelector selector, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) =>
        Text(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to extract plain text from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<string> TryText(PdfPageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts text per page in the order resolved by a document-relative selector.</summary>
    public IReadOnlyList<string> TextByPage(PdfPageSelector selector, PdfReadOptions? readOptions = null) =>
        TextByPage(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract text per page resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts Markdown from pages resolved by a document-relative selector.</summary>
    public string Markdown(PdfPageSelector selector, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? readOptions = null) =>
        Markdown(ResolveSelector(selector, readOptions), options, markdownOptions, readOptions);

    /// <summary>Attempts to extract Markdown from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<string> TryMarkdown(PdfPageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(selector, layoutOptions, markdownOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Builds the logical model for pages resolved by a document-relative selector.</summary>
    public PdfLogicalDocument Logical(PdfPageSelector selector, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) =>
        Logical(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to build the logical model for pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(PdfPageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts logical text blocks from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks(PdfPageSelector selector, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) =>
        TextBlocks(ResolveSelector(selector, readOptions), options, readOptions);

    /// <summary>Attempts to extract logical text blocks from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> TryTextBlocks(PdfPageSelector selector, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract logical text blocks", PdfPreflightCapability.ReadLogicalObjects, () => TextBlocks(selector, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts images from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfExtractedImage> Images(PdfPageSelector selector, PdfReadOptions? readOptions = null) =>
        Images(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract images from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> TryImages(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, () => Images(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Extracts image placements from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements(PdfPageSelector selector, PdfReadOptions? readOptions = null) =>
        ImagePlacements(ResolveSelector(selector, readOptions), readOptions);

    /// <summary>Attempts to extract image placements from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<PdfImagePlacement>> TryImagePlacements(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation("Extract image placements", PdfPreflightCapability.ExtractImages, () => ImagePlacements(selector, options), ResolveReadOptions(options));
    }

    /// <summary>Inspects unique fonts declared by pages resolved by a document-relative selector.</summary>
    public PdfFontInventory Fonts(
        PdfPageSelector selector,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? readOptions = null) =>
        Fonts(ResolveSelector(selector, readOptions), inspectionOptions, readOptions);

    /// <summary>Attempts to inspect fonts on pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<PdfFontInventory> TryFonts(
        PdfPageSelector selector,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Inspect fonts",
            PdfPreflightCapability.ReadLogicalObjects,
            () => Fonts(selector, inspectionOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Recovers conservative paragraph continuation groups from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalParagraphContinuationGroup> ParagraphContinuations(
        PdfPageSelector selector,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) =>
        ParagraphContinuations(ResolveSelector(selector, readOptions), continuationOptions, layoutOptions, readOptions);

    /// <summary>Attempts to recover paragraph continuation groups from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> TryParagraphContinuations(
        PdfPageSelector selector,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Recover paragraph continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => ParagraphContinuations(selector, continuationOptions, layoutOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Recovers bounded table continuation groups from pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> TableContinuations(
        PdfPageSelector selector,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) =>
        TableContinuations(ResolveSelector(selector, readOptions), continuationOptions, layoutOptions, readOptions);

    /// <summary>Attempts to recover table continuation groups from pages resolved by a document-relative selector.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> TryTableContinuations(
        PdfPageSelector selector,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryOperation(
            "Recover table continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => TableContinuations(selector, continuationOptions, layoutOptions, options),
            ResolveReadOptions(options));
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        PdfPageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(selector, nameof(selector));
        return PdfPageImageRenderer.RenderPages(
            _document.GetBytesForOperation,
            selector,
            options,
            ResolveReadOptions(readOptions),
            cancellationToken);
    }

    /// <summary>Runs the understanding pipeline for pages resolved by a document-relative selector.</summary>
    public PdfUnderstandingResult Understand(
        PdfPageSelector selector,
        PdfUnderstandingPipelineOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadDocument document = ReadDocument(readOptions);
        return new PdfUnderstandingPipeline(options).Run(document, selector, cancellationToken);
    }

    private PdfPageSelection ResolveSelector(PdfPageSelector selector, PdfReadOptions? readOptions) {
        Guard.NotNull(selector, nameof(selector));
        int pageCount = _document.Inspect(ResolveReadOptions(readOptions)).PageCount;
        if (pageCount < 1) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        return selector.ResolveSelection(pageCount);
    }
}
