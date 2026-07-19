namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent readback operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentReader {
    private readonly PdfDocument _document;

    internal PdfDocumentReader(PdfDocument document) {
        _document = document;
    }

    private PdfReadOptions ResolveReadOptions(PdfReadOptions? readOptions) {
        return readOptions ?? _document.ReadOptions;
    }

    private PdfReadDocument ReadDocument(PdfReadOptions? readOptions = null) =>
        _document.GetReadDocument(ResolveReadOptions(readOptions));

    /// <summary>
    /// Extracts plain text from all pages.
    /// </summary>
    public string Text(PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        PdfReadDocument document = ReadDocument(readOptions);
        return options is null ? document.ExtractText() : document.ExtractTextWithColumns(options);
    }

    /// <summary>
    /// Attempts to extract plain text from all pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryText(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts plain text from selected pages and concatenates them with blank lines.
    /// </summary>
    public string Text(PdfPageSelection selection, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfTextExtractor.ExtractAllTextByPageRanges(ReadDocument(readOptions), options, selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract plain text from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryText(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(selection, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts plain text from selected pages and concatenates them with blank lines.
    /// </summary>
    public string Text(string pageRanges, PdfTextLayoutOptions? options = null) {
        return Text(PdfPageSelection.Parse(pageRanges), options);
    }

    /// <summary>
    /// Attempts to extract plain text from selected pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryText(string pageRanges, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text", PdfPreflightCapability.ExtractText, () => Text(PdfPageSelection.Parse(pageRanges), layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts plain text for each page.
    /// </summary>
    public IReadOnlyList<string> TextByPage(PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        PdfReadDocument document = ReadDocument(readOptions);
        return options is null
            ? PdfTextExtractor.ExtractTextByPage(document)
            : PdfTextExtractor.ExtractTextByPage(document, options);
    }

    /// <summary>
    /// Attempts to extract plain text for each page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts plain text for selected pages in caller order.
    /// </summary>
    public IReadOnlyList<string> TextByPage(PdfPageSelection selection, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfTextExtractor.ExtractTextByPageRanges(ReadDocument(readOptions), selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract plain text for selected pages in caller order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(selection, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts plain text for selected pages in caller order.
    /// </summary>
    public IReadOnlyList<string> TextByPage(string pageRanges) {
        return TextByPage(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to extract plain text for selected pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<string>> TryTextByPage(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract text by page", PdfPreflightCapability.ExtractText, () => TextByPage(PdfPageSelection.Parse(pageRanges), options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts Markdown from the logical readback model.
    /// </summary>
    public string Markdown(PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? readOptions = null) {
        return PdfLogicalDocument.From(ReadDocument(readOptions), options).ToMarkdown(markdownOptions);
    }

    /// <summary>
    /// Attempts to extract Markdown from the logical readback model, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryMarkdown(PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(layoutOptions, markdownOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts Markdown from selected pages in the logical readback model.
    /// </summary>
    public string Markdown(PdfPageSelection selection, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfLogicalDocument
            .FromPageRanges(ReadDocument(readOptions), options, selection.ToRanges())
            .ToMarkdown(markdownOptions);
    }

    /// <summary>
    /// Attempts to extract Markdown from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryMarkdown(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(selection, layoutOptions, markdownOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts Markdown from selected pages in the logical readback model.
    /// </summary>
    public string Markdown(string pageRanges, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        return Markdown(PdfPageSelection.Parse(pageRanges), options, markdownOptions);
    }

    /// <summary>
    /// Attempts to extract Markdown from selected pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<string> TryMarkdown(string pageRanges, PdfTextLayoutOptions? layoutOptions = null, PdfLogicalMarkdownOptions? markdownOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract Markdown", PdfPreflightCapability.ReadLogicalObjects, () => Markdown(PdfPageSelection.Parse(pageRanges), layoutOptions, markdownOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Builds the logical document model.
    /// </summary>
    public PdfLogicalDocument Logical(PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        return PdfLogicalDocument.From(ReadDocument(readOptions), options);
    }

    /// <summary>
    /// Attempts to build the logical document model, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Builds the logical document model for selected pages.
    /// </summary>
    public PdfLogicalDocument Logical(PdfPageSelection selection, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfLogicalDocument.FromPageRanges(ReadDocument(readOptions), options, selection.ToRanges());
    }

    /// <summary>
    /// Attempts to build the logical document model for selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(selection, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Builds the logical document model for selected pages.
    /// </summary>
    public PdfLogicalDocument Logical(string pageRanges, PdfTextLayoutOptions? options = null) {
        return Logical(PdfPageSelection.Parse(pageRanges), options);
    }

    /// <summary>
    /// Attempts to build the logical document model for selected pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfLogicalDocument> TryLogical(string pageRanges, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Read logical document", PdfPreflightCapability.ReadLogicalObjects, () => Logical(PdfPageSelection.Parse(pageRanges), layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts line-level logical text blocks from all pages.
    /// </summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks(PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        return Logical(options, readOptions).TextBlocks;
    }

    /// <summary>
    /// Attempts to extract line-level logical text blocks from all pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> TryTextBlocks(PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract logical text blocks", PdfPreflightCapability.ReadLogicalObjects, () => TextBlocks(layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts line-level logical text blocks from selected pages.
    /// </summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks(PdfPageSelection selection, PdfTextLayoutOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return Logical(selection, options, readOptions).TextBlocks;
    }

    /// <summary>
    /// Attempts to extract line-level logical text blocks from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> TryTextBlocks(PdfPageSelection selection, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract logical text blocks", PdfPreflightCapability.ReadLogicalObjects, () => TextBlocks(selection, layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts line-level logical text blocks from selected pages.
    /// </summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks(string pageRanges, PdfTextLayoutOptions? options = null) {
        return TextBlocks(PdfPageSelection.Parse(pageRanges), options);
    }

    /// <summary>
    /// Attempts to extract line-level logical text blocks from selected pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> TryTextBlocks(string pageRanges, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract logical text blocks", PdfPreflightCapability.ReadLogicalObjects, () => TextBlocks(PdfPageSelection.Parse(pageRanges), layoutOptions, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads top-level document outline/bookmark entries from the logical model.
    /// </summary>
    public IReadOnlyList<PdfOutlineItem> Outlines(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).Outlines;
    }

    /// <summary>
    /// Attempts to read top-level document outline/bookmark entries, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOutlineItem>> TryOutlines(PdfReadOptions? options = null) {
        return _document.TryOperation("Read outlines", PdfPreflightCapability.ReadLogicalObjects, () => Outlines(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-label rules discovered from the document catalog.
    /// </summary>
    public IReadOnlyList<PdfPageLabel> PageLabels(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).PageLabels;
    }

    /// <summary>
    /// Attempts to read page-label rules, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageLabel>> TryPageLabels(PdfReadOptions? options = null) {
        return _document.TryOperation("Read page labels", PdfPreflightCapability.ReadLogicalObjects, () => PageLabels(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads named destinations discovered from the document catalog.
    /// </summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).NamedDestinations;
    }

    /// <summary>
    /// Attempts to read named destinations, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfNamedDestination>> TryNamedDestinations(PdfReadOptions? options = null) {
        return _document.TryOperation("Read named destinations", PdfPreflightCapability.ReadLogicalObjects, () => NamedDestinations(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads a simple document open action from the document catalog, when supported.
    /// </summary>
    public PdfDocumentOpenAction? OpenAction(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).OpenAction;
    }

    /// <summary>
    /// Reads simple viewer preference entries from the document catalog, when supported.
    /// </summary>
    public PdfViewerPreferences? ViewerPreferences(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).ViewerPreferences;
    }

    /// <summary>
    /// Reads the catalog page mode name, for example UseOutlines or FullScreen, when present.
    /// </summary>
    public string? CatalogPageMode(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).CatalogPageMode;
    }

    /// <summary>
    /// Reads the catalog page layout name, for example SinglePage or TwoColumnLeft, when present.
    /// </summary>
    public string? CatalogPageLayout(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).CatalogPageLayout;
    }

    /// <summary>
    /// Reads the catalog PDF version override, for example 1.7, when present.
    /// </summary>
    public string? CatalogVersion(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).CatalogVersion;
    }

    /// <summary>
    /// Reads the catalog language tag, for example en-US or pl-PL, when present.
    /// </summary>
    public string? CatalogLanguage(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).CatalogLanguage;
    }

    /// <summary>
    /// Extracts URI, named-destination, direct-destination, named-action, and remote GoTo link annotations from the logical model.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> Links(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).Links;
    }

    /// <summary>
    /// Attempts to extract logical link annotations, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinks(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => Links(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts logical URI link annotations for a URI action target.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> LinksByUri(string uri, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetLinksByUri(uri);
    }

    /// <summary>
    /// Attempts to extract logical URI link annotations for a URI action target, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinksByUri(string uri, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => LinksByUri(uri, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts logical internal link annotations for a named destination.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> LinksByDestinationName(string destinationName, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetLinksByDestinationName(destinationName);
    }

    /// <summary>
    /// Attempts to extract logical internal link annotations for a named destination, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinksByDestinationName(string destinationName, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => LinksByDestinationName(destinationName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts logical internal direct-destination link annotations for a one-based destination page number.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> LinksByDestinationPageNumber(int pageNumber, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetLinksByDestinationPageNumber(pageNumber);
    }

    /// <summary>
    /// Attempts to extract logical internal direct-destination link annotations for a one-based destination page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinksByDestinationPageNumber(int pageNumber, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => LinksByDestinationPageNumber(pageNumber, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts logical named-action link annotations for a viewer action name.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> LinksByNamedAction(string namedAction, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetLinksByNamedAction(namedAction);
    }

    /// <summary>
    /// Attempts to extract logical named-action link annotations for a viewer action name, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinksByNamedAction(string namedAction, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => LinksByNamedAction(namedAction, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts logical remote GoTo link annotations for a target file.
    /// </summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> LinksByRemoteFile(string remoteFile, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetLinksByRemoteFile(remoteFile);
    }

    /// <summary>
    /// Attempts to extract logical remote GoTo link annotations for a target file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalLinkAnnotation>> TryLinksByRemoteFile(string remoteFile, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract links", PdfPreflightCapability.ReadLogicalObjects, () => LinksByRemoteFile(remoteFile, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads generic page annotations in document order.
    /// </summary>
    public IReadOnlyList<PdfAnnotation> Annotations(PdfReadOptions? readOptions = null) {
        return _document.Inspect(ResolveReadOptions(readOptions)).Annotations;
    }

    /// <summary>
    /// Attempts to read generic page annotations, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAnnotation>> TryAnnotations(PdfReadOptions? options = null) {
        return _document.TryOperation("Read annotations", PdfPreflightCapability.ReadLogicalObjects, () => Annotations(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads generic page annotations with a matching PDF annotation subtype name.
    /// </summary>
    public IReadOnlyList<PdfAnnotation> AnnotationsBySubtype(string subtype, PdfReadOptions? readOptions = null) {
        return _document.Inspect(ResolveReadOptions(readOptions)).GetAnnotationsBySubtype(subtype);
    }

    /// <summary>
    /// Attempts to read generic page annotations with a matching PDF annotation subtype name, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAnnotation>> TryAnnotationsBySubtype(string subtype, PdfReadOptions? options = null) {
        return _document.TryOperation("Read annotations", PdfPreflightCapability.ReadLogicalObjects, () => AnnotationsBySubtype(subtype, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads generic page annotations with a matching primary or additional action type.
    /// </summary>
    public IReadOnlyList<PdfAnnotation> AnnotationsByActionType(string actionType, PdfReadOptions? readOptions = null) {
        return _document.Inspect(ResolveReadOptions(readOptions)).GetAnnotationsByActionType(actionType);
    }

    /// <summary>
    /// Attempts to read generic page annotations with a matching primary or additional action type, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAnnotation>> TryAnnotationsByActionType(string actionType, PdfReadOptions? options = null) {
        return _document.TryOperation("Read annotations", PdfPreflightCapability.ReadLogicalObjects, () => AnnotationsByActionType(actionType, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts simple AcroForm fields from the logical readback model.
    /// </summary>
    public IReadOnlyList<PdfFormField> FormFields(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).FormFields;
    }

    /// <summary>
    /// Attempts to extract simple AcroForm fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfFormField>> TryFormFields(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form fields", PdfPreflightCapability.ReadLogicalObjects, () => FormFields(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Returns the simple AcroForm field with the requested fully qualified field name, when present.
    /// </summary>
    public PdfFormField? FormField(string fieldName, PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(fieldName, nameof(fieldName));
        return Logical(readOptions: readOptions).TryGetFormField(fieldName, out PdfFormField? field)
            ? field
            : null;
    }

    /// <summary>
    /// Extracts simple AcroForm fields matching the requested fully qualified field name.
    /// </summary>
    public IReadOnlyList<PdfFormField> FormFields(string fieldName, PdfReadOptions? readOptions = null) {
        PdfFormField? field = FormField(fieldName, readOptions);
        return field is null
            ? Array.Empty<PdfFormField>()
            : new[] { field };
    }

    /// <summary>
    /// Attempts to extract simple AcroForm fields matching the requested fully qualified field name, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfFormField>> TryFormFields(string fieldName, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form fields", PdfPreflightCapability.ReadLogicalObjects, () => FormFields(fieldName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts simple AcroForm fields for the requested common field kind.
    /// </summary>
    public IReadOnlyList<PdfFormField> FormFields(PdfFormFieldKind kind, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetFormFields(kind);
    }

    /// <summary>
    /// Attempts to extract simple AcroForm fields for the requested common field kind, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfFormField>> TryFormFields(PdfFormFieldKind kind, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form fields", PdfPreflightCapability.ReadLogicalObjects, () => FormFields(kind, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts simple AcroForm fields represented by widgets on a one-based page number.
    /// </summary>
    public IReadOnlyList<PdfFormField> FormFields(int pageNumber, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetFormFields(pageNumber);
    }

    /// <summary>
    /// Attempts to extract simple AcroForm fields represented by widgets on a one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfFormField>> TryFormFields(int pageNumber, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form fields", PdfPreflightCapability.ReadLogicalObjects, () => FormFields(pageNumber, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts AcroForm widget annotations with page geometry from the logical readback model.
    /// </summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets(PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).FormWidgets;
    }

    /// <summary>
    /// Attempts to extract AcroForm widget annotations with page geometry, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalFormWidget>> TryFormWidgets(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form widgets", PdfPreflightCapability.ReadLogicalObjects, () => FormWidgets(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts AcroForm widget annotations for the requested fully qualified field name.
    /// </summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets(string fieldName, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetFormWidgets(fieldName);
    }

    /// <summary>
    /// Attempts to extract AcroForm widget annotations for the requested fully qualified field name, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalFormWidget>> TryFormWidgets(string fieldName, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form widgets", PdfPreflightCapability.ReadLogicalObjects, () => FormWidgets(fieldName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts AcroForm widget annotations from a one-based page number.
    /// </summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets(int pageNumber, PdfReadOptions? readOptions = null) {
        return Logical(readOptions: readOptions).GetFormWidgets(pageNumber);
    }

    /// <summary>
    /// Attempts to extract AcroForm widget annotations from a one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalFormWidget>> TryFormWidgets(int pageNumber, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract form widgets", PdfPreflightCapability.ReadLogicalObjects, () => FormWidgets(pageNumber, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObjects.
    /// </summary>
    public IReadOnlyList<PdfExtractedImage> Images(PdfReadOptions? readOptions = null) {
        return ReadDocument(readOptions).ExtractImages();
    }

    /// <summary>
    /// Attempts to extract image XObjects, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> TryImages(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, () => Images(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObjects from selected pages in caller order.
    /// </summary>
    public IReadOnlyList<PdfExtractedImage> Images(PdfPageSelection selection, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfImageExtractor.ExtractImagesByPageRanges(ReadDocument(readOptions), selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract image XObjects from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> TryImages(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, () => Images(selection, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObjects from pages described by comma- or semicolon-separated inclusive page ranges.
    /// </summary>
    public IReadOnlyList<PdfExtractedImage> Images(string pageRanges, PdfReadOptions? readOptions = null) {
        return Images(PdfPageSelection.Parse(pageRanges), readOptions);
    }

    /// <summary>Extracts all readable image XObjects to deterministic files.</summary>
    public IReadOnlyList<string> SaveImages(
        string outputDirectory,
        string baseName = "image",
        PdfReadOptions? readOptions = null) {
        return PdfImageExtractor.WriteImages(Images(readOptions), outputDirectory, baseName);
    }

    /// <summary>Extracts image XObjects from selected pages to deterministic files.</summary>
    public IReadOnlyList<string> SaveImages(
        string outputDirectory,
        PdfPageSelection selection,
        string baseName = "image",
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfImageExtractor.WriteImages(Images(selection, readOptions), outputDirectory, baseName);
    }

    /// <summary>
    /// Attempts to extract image XObjects from pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> TryImages(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract images", PdfPreflightCapability.ExtractImages, () => Images(PdfPageSelection.Parse(pageRanges), options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObject placement invocations with page geometry.
    /// </summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements(PdfReadOptions? readOptions = null) {
        return PdfImageExtractor.ExtractImagePlacements(ReadDocument(readOptions));
    }

    /// <summary>
    /// Attempts to extract image XObject placement invocations with page geometry, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfImagePlacement>> TryImagePlacements(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract image placements", PdfPreflightCapability.ExtractImages, () => ImagePlacements(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObject placement invocations from selected pages in caller order.
    /// </summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements(PdfPageSelection selection, PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfImageExtractor.ExtractImagePlacementsByPageRanges(ReadDocument(readOptions), selection.ToRanges());
    }

    /// <summary>
    /// Attempts to extract image XObject placement invocations from selected pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfImagePlacement>> TryImagePlacements(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract image placements", PdfPreflightCapability.ExtractImages, () => ImagePlacements(selection, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts image XObject placement invocations from pages described by comma- or semicolon-separated inclusive page ranges.
    /// </summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements(string pageRanges, PdfReadOptions? readOptions = null) {
        return ImagePlacements(PdfPageSelection.Parse(pageRanges), readOptions);
    }

    /// <summary>
    /// Attempts to extract image XObject placement invocations from pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfImagePlacement>> TryImagePlacements(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryOperation("Extract image placements", PdfPreflightCapability.ExtractImages, () => ImagePlacements(PdfPageSelection.Parse(pageRanges), options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Extracts embedded-file attachments.
    /// </summary>
    public IReadOnlyList<PdfExtractedAttachment> Attachments(PdfReadOptions? readOptions = null) {
        return ReadDocument(readOptions).ExtractAttachments();
    }

    /// <summary>
    /// Attempts to extract embedded-file attachments, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedAttachment>> TryAttachments(PdfReadOptions? options = null) {
        return _document.TryOperation("Extract attachments", PdfPreflightCapability.ExtractAttachments, () => Attachments(options), ResolveReadOptions(options));
    }
}
