namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent active-action readback operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentReader {
    /// <summary>
    /// Reads catalog-level actions discovered from supported catalog slots and name trees.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).CatalogActions;
    }

    /// <summary>
    /// Attempts to read catalog-level actions, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActions(PdfReadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActions(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads catalog-level actions with a matching PDF action type.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActionsByActionType(string actionType, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetCatalogActionsByActionType(actionType);
    }

    /// <summary>
    /// Attempts to read catalog-level actions with a matching PDF action type, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActionsByActionType(string actionType, PdfReadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActionsByActionType(actionType, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads catalog-level actions from a matching catalog source.
    /// </summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActionsBySource(string source, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetCatalogActionsBySource(source);
    }

    /// <summary>
    /// Attempts to read catalog-level actions from a matching catalog source, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfCatalogAction>> TryCatalogActionsBySource(string source, PdfReadOptions? options = null) {
        return _document.TryOperation("Read catalog actions", PdfPreflightCapability.ReadLogicalObjects, () => CatalogActionsBySource(source, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions discovered from page dictionaries.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActions(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).PageActions;
    }

    /// <summary>
    /// Attempts to read page-level additional actions, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActions(PdfReadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActions(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions for a one-based page number.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActions(int pageNumber, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActions(pageNumber);
    }

    /// <summary>
    /// Attempts to read page-level additional actions for a one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActions(int pageNumber, PdfReadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActions(pageNumber, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching PDF action type.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByActionType(string actionType, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByActionType(actionType);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching PDF action type, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByActionType(string actionType, PdfReadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByActionType(actionType, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching page /AA trigger key.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByTriggerName(string triggerName, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByTriggerName(triggerName);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching page /AA trigger key, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByTriggerName(string triggerName, PdfReadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByTriggerName(triggerName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page-level additional actions with a matching stable action path.
    /// </summary>
    public IReadOnlyList<PdfPageAction> PageActionsByActionPath(string actionPath, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetPageActionsByActionPath(actionPath);
    }

    /// <summary>
    /// Attempts to read page-level additional actions with a matching stable action path, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageAction>> TryPageActionsByActionPath(string actionPath, PdfReadOptions? options = null) {
        return _document.TryOperation("Read page actions", PdfPreflightCapability.ReadLogicalObjects, () => PageActionsByActionPath(actionPath, options), ResolveReadOptions(options));
    }
}
