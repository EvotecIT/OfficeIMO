namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent readback operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed class PdfDocumentReader {
    private readonly PdfDocument _document;

    internal PdfDocumentReader(PdfDocument document) {
        _document = document;
    }

    /// <summary>
    /// Extracts plain text from all pages.
    /// </summary>
    public string Text(PdfTextLayoutOptions? options = null) {
        return PdfTextExtractor.ExtractAllText(_document.Snapshot(), options);
    }

    /// <summary>
    /// Attempts to extract plain text from all pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryText(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(layoutOptions), options);
    }

    /// <summary>
    /// Extracts plain text from selected pages and concatenates them with blank lines.
    /// </summary>
    public string Text(PdfPageSelection selection, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfTextExtractor.ExtractAllTextByPageRanges(_document.Snapshot(), options, selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract plain text from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryText(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(selection, layoutOptions), options);
    }

    /// <summary>
    /// Extracts plain text from selected pages and concatenates them with blank lines.
    /// </summary>
    public string Text(string pageRanges, PdfTextLayoutOptions? options = null) {
        return Text(PdfPageSelection.Parse(pageRanges), options);
    }

    /// <summary>
    /// Extracts plain text for each page.
    /// </summary>
    public IReadOnlyList<string> TextByPage(PdfTextLayoutOptions? options = null) {
        return PdfTextExtractor.ExtractTextByPage(_document.Snapshot(), options);
    }

    /// <summary>
    /// Attempts to extract plain text for each page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(layoutOptions), options);
    }

    /// <summary>
    /// Extracts plain text for selected pages in caller order.
    /// </summary>
    public IReadOnlyList<string> TextByPage(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfTextExtractor.ExtractTextByPageRanges(_document.Snapshot(), selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract plain text for selected pages in caller order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(selection), options);
    }

    /// <summary>
    /// Extracts plain text for selected pages in caller order.
    /// </summary>
    public IReadOnlyList<string> TextByPage(string pageRanges) {
        return TextByPage(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Extracts Markdown from the logical readback model.
    /// </summary>
    public string Markdown(PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        return PdfTextExtractor.ExtractMarkdown(_document.Snapshot(), options, markdownOptions);
    }

    /// <summary>
    /// Attempts to extract Markdown from the logical readback model, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryMarkdown(PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(layoutOptions, markdownOptions), options);
    }

    /// <summary>
    /// Extracts Markdown from selected pages in the logical readback model.
    /// </summary>
    public string Markdown(PdfPageSelection selection, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(_document.Snapshot(), options, markdownOptions, selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract Markdown from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryMarkdown(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(selection, layoutOptions, markdownOptions), options);
    }

    /// <summary>
    /// Extracts Markdown from selected pages in the logical readback model.
    /// </summary>
    public string Markdown(string pageRanges, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        return Markdown(PdfPageSelection.Parse(pageRanges), options, markdownOptions);
    }

    /// <summary>
    /// Builds the logical document model.
    /// </summary>
    public PdfLogicalDocument Logical(PdfTextLayoutOptions? options = null) {
        return PdfLogicalDocument.Load(_document.Snapshot(), options);
    }

    /// <summary>
    /// Attempts to build the logical document model, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(layoutOptions), options);
    }

    /// <summary>
    /// Builds the logical document model for selected pages.
    /// </summary>
    public PdfLogicalDocument Logical(PdfPageSelection selection, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfLogicalDocument.LoadPageRanges(_document.Snapshot(), options, selection.ToRanges());
    }

    /// <summary>
    /// Attempts to build the logical document model for selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(selection, layoutOptions), options);
    }

    /// <summary>
    /// Builds the logical document model for selected pages.
    /// </summary>
    public PdfLogicalDocument Logical(string pageRanges, PdfTextLayoutOptions? options = null) {
        return Logical(PdfPageSelection.Parse(pageRanges), options);
    }

    /// <summary>
    /// Extracts image XObjects.
    /// </summary>
    public IReadOnlyList<PdfExtractedImage> Images() {
        return PdfImageExtractor.ExtractImages(_document.Snapshot());
    }

    /// <summary>
    /// Attempts to extract image XObjects, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> TryImages(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, Images, options);
    }

    /// <summary>
    /// Extracts embedded-file attachments.
    /// </summary>
    public IReadOnlyList<PdfExtractedAttachment> Attachments() {
        return PdfAttachmentExtractor.ExtractAttachments(_document.Snapshot());
    }
}
