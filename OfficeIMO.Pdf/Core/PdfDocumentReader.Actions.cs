namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent active-action readback operations for a <see cref="PdfDocument"/>.
/// </summary>
internal sealed partial class PdfDocumentReader {
    /// <summary>Reads named document-level JavaScript actions and their exact source text.</summary>
    public IReadOnlyList<PdfJavaScript> JavaScripts(PdfLoadOptions? readOptions = null) =>
        ReadDocument(readOptions).JavaScripts;

    /// <summary>Attempts to read named document-level JavaScript actions, returning diagnostics when blocked or failed.</summary>
    public PdfOperationResult<IReadOnlyList<PdfJavaScript>> TryJavaScripts(PdfLoadOptions? options = null) =>
        _document.TryOperation("Read document JavaScript", PdfPreflightCapability.ReadLogicalObjects, () => JavaScripts(options), ResolveReadOptions(options));

    /// <summary>
    /// Reads catalog-level actions discovered from supported catalog slots and name trees.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions(PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).CatalogActions;
    }

    /// <summary>
    /// Attempts to read catalog-level actions, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActions(PdfLoadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActions(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads catalog-level actions with a matching PDF action type.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActionsByActionType(string actionType, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetCatalogActionsByActionType(actionType);
    }

    /// <summary>
    /// Attempts to read catalog-level actions with a matching PDF action type, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActionsByActionType(string actionType, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActionsByActionType(actionType, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads catalog-level actions from a matching catalog source.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActionsBySource(string source, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetCatalogActionsBySource(source);
    }

    /// <summary>
    /// Attempts to read catalog-level actions from a matching catalog source, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActionsBySource(string source, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActionsBySource(source, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions discovered from page dictionaries.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActions(PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).PageActions;
    }

    /// <summary>
    /// Attempts to read page-level additional actions, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActions(PdfLoadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActions(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions for a one-based page number.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActions(int pageNumber, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActions(pageNumber);
    }

    /// <summary>
    /// Attempts to read page-level additional actions for a one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActions(int pageNumber, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActions(pageNumber, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching PDF action type.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByActionType(string actionType, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByActionType(actionType);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching PDF action type, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByActionType(string actionType, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByActionType(actionType, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching page /AA trigger key.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByTriggerName(string triggerName, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByTriggerName(triggerName);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching page /AA trigger key, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByTriggerName(string triggerName, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByTriggerName(triggerName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching stable action path.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByActionPath(string actionPath, PdfLoadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByActionPath(actionPath);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching stable action path, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByActionPath(string actionPath, PdfLoadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByActionPath(actionPath, options), ResolveReadOptions(options));
    }
}
