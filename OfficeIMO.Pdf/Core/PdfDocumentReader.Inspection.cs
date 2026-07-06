namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent inspection readback operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentReader {
    /// <summary>
    /// Reads document-level inspection information through the shared inspector model.
    /// </summary>
    public PdfDocumentInfo DocumentInfo(PdfReadOptions? readOptions = null) {
        return _document.Inspect(ResolveReadOptions(readOptions));
    }

    /// <summary>
    /// Attempts to read document-level inspection information, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocumentInfo> TryDocumentInfo(PdfReadOptions? options = null) {
        return _document.TryOperation("Read document info", PdfPreflightCapability.ReadLogicalObjects, () => DocumentInfo(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads PDF document metadata from the Info dictionary.
    /// </summary>
    public PdfMetadata Metadata(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).Metadata;
    }

    /// <summary>
    /// Attempts to read PDF document metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfMetadata> TryMetadata(PdfReadOptions? options = null) {
        return _document.TryOperation("Read metadata", PdfPreflightCapability.ReadLogicalObjects, () => Metadata(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads security, signature, and revision markers.
    /// </summary>
    public PdfDocumentSecurityInfo Security(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).Security;
    }

    /// <summary>
    /// Attempts to read security, signature, and revision markers, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocumentSecurityInfo> TrySecurity(PdfReadOptions? options = null) {
        return _document.TryOperation("Read security markers", PdfPreflightCapability.ReadLogicalObjects, () => Security(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page geometry and page-level metadata for all pages.
    /// </summary>
    public IReadOnlyList<PdfPageInfo> Pages(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).Pages;
    }

    /// <summary>
    /// Attempts to read page geometry and page-level metadata for all pages, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfPageInfo>> TryPages(PdfReadOptions? options = null) {
        return _document.TryOperation("Read page info", PdfPreflightCapability.ReadLogicalObjects, () => Pages(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads page geometry and page-level metadata for a one-based page number, or null when the page does not exist.
    /// </summary>
    public PdfPageInfo? Page(int pageNumber, PdfReadOptions? readOptions = null) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        IReadOnlyList<PdfPageInfo> pages = Pages(readOptions);
        for (int i = 0; i < pages.Count; i++) {
            if (pages[i].PageNumber == pageNumber) {
                return pages[i];
            }
        }

        return null;
    }

    /// <summary>
    /// Reads the PDF header version, for example 1.7, when present.
    /// </summary>
    public string? HeaderVersion(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).HeaderVersion;
    }

    /// <summary>
    /// Reads the effective PDF version inferred from catalog override or file header version.
    /// </summary>
    public string? EffectiveVersion(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).EffectiveVersion;
    }

    /// <summary>
    /// Returns true when the effective PDF version is PDF 2.0 or later.
    /// </summary>
    public bool IsPdf20OrLater(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).IsPdf20OrLater;
    }
}
