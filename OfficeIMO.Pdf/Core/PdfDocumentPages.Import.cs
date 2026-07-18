namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>
    /// Appends all pages from another loaded or generated PDF to this document.
    /// </summary>
    public PdfDocument Append(PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Append(sourceDocument.GetBytesForOperation(), importOptions);
    }

    /// <summary>
    /// Appends selected one-based pages from another loaded or generated PDF to this document.
    /// </summary>
    public PdfDocument Append(PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Append(sourceDocument.GetBytesForOperation(), sourceSelection, importOptions);
    }

    /// <summary>
    /// Appends all pages from source PDF bytes to this document.
    /// </summary>
    public PdfDocument Append(byte[] sourcePdf, PdfPageImportOptions? importOptions = null) {
        return Import(sourcePdf, ImportPlacement.Append, insertBeforePageNumber: null, sourcePageNumbers: Array.Empty<int>(), importOptions);
    }

    /// <summary>
    /// Appends selected one-based pages from source PDF bytes to this document.
    /// </summary>
    public PdfDocument Append(byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return Import(sourcePdf, ImportPlacement.Append, insertBeforePageNumber: null, GetSelectedSourcePages(sourcePdf, sourceSelection), importOptions);
    }

    /// <summary>
    /// Appends all pages from a readable source PDF stream to this document.
    /// </summary>
    public PdfDocument Append(Stream sourceStream, PdfPageImportOptions? importOptions = null) {
        return Append(ReadSourceStream(sourceStream), importOptions);
    }

    /// <summary>
    /// Appends selected one-based pages from a readable source PDF stream to this document.
    /// </summary>
    public PdfDocument Append(Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        return Append(ReadSourceStream(sourceStream), sourceSelection, importOptions);
    }

    /// <summary>
    /// Appends all pages from a source PDF file to this document.
    /// </summary>
    public PdfDocument Append(string sourcePath, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Append(File.ReadAllBytes(sourcePath), importOptions);
    }

    /// <summary>
    /// Appends selected one-based pages from a source PDF file to this document.
    /// </summary>
    public PdfDocument Append(string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Append(File.ReadAllBytes(sourcePath), sourceSelection, importOptions);
    }

    /// <summary>
    /// Attempts to append pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourceDocument, importOptions), options);
    }

    /// <summary>
    /// Attempts to append selected pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourceDocument, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to append source PDF bytes, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(byte[] sourcePdf, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourcePdf, importOptions), options);
    }

    /// <summary>
    /// Attempts to append selected pages from source PDF bytes, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourcePdf, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to append a readable source PDF stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(Stream sourceStream, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourceStream, importOptions), options);
    }

    /// <summary>
    /// Attempts to append selected pages from a readable source PDF stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourceStream, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to append a source PDF file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(string sourcePath, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourcePath, importOptions), options);
    }

    /// <summary>
    /// Attempts to append selected pages from a source PDF file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppend(string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Append pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Append(sourcePath, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Prepends all pages from another loaded or generated PDF before this document.
    /// </summary>
    public PdfDocument Prepend(PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Prepend(sourceDocument.GetBytesForOperation(), importOptions);
    }

    /// <summary>
    /// Prepends selected one-based pages from another loaded or generated PDF before this document.
    /// </summary>
    public PdfDocument Prepend(PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Prepend(sourceDocument.GetBytesForOperation(), sourceSelection, importOptions);
    }

    /// <summary>
    /// Prepends all pages from source PDF bytes before this document.
    /// </summary>
    public PdfDocument Prepend(byte[] sourcePdf, PdfPageImportOptions? importOptions = null) {
        return Import(sourcePdf, ImportPlacement.Prepend, insertBeforePageNumber: null, sourcePageNumbers: Array.Empty<int>(), importOptions);
    }

    /// <summary>
    /// Prepends selected one-based pages from source PDF bytes before this document.
    /// </summary>
    public PdfDocument Prepend(byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return Import(sourcePdf, ImportPlacement.Prepend, insertBeforePageNumber: null, GetSelectedSourcePages(sourcePdf, sourceSelection), importOptions);
    }

    /// <summary>
    /// Prepends all pages from a readable source PDF stream before this document.
    /// </summary>
    public PdfDocument Prepend(Stream sourceStream, PdfPageImportOptions? importOptions = null) {
        return Prepend(ReadSourceStream(sourceStream), importOptions);
    }

    /// <summary>
    /// Prepends selected one-based pages from a readable source PDF stream before this document.
    /// </summary>
    public PdfDocument Prepend(Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        return Prepend(ReadSourceStream(sourceStream), sourceSelection, importOptions);
    }

    /// <summary>
    /// Prepends all pages from a source PDF file before this document.
    /// </summary>
    public PdfDocument Prepend(string sourcePath, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Prepend(File.ReadAllBytes(sourcePath), importOptions);
    }

    /// <summary>
    /// Prepends selected one-based pages from a source PDF file before this document.
    /// </summary>
    public PdfDocument Prepend(string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Prepend(File.ReadAllBytes(sourcePath), sourceSelection, importOptions);
    }

    /// <summary>
    /// Attempts to prepend pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourceDocument, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend selected pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourceDocument, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend source PDF bytes, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(byte[] sourcePdf, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourcePdf, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend selected pages from source PDF bytes, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourcePdf, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend a readable source PDF stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(Stream sourceStream, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourceStream, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend selected pages from a readable source PDF stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourceStream, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend a source PDF file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(string sourcePath, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourcePath, importOptions), options);
    }

    /// <summary>
    /// Attempts to prepend selected pages from a source PDF file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryPrepend(string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Prepend pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Prepend(sourcePath, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Inserts all pages from another loaded or generated PDF before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Insert(insertBeforePageNumber, sourceDocument.GetBytesForOperation(), importOptions);
    }

    /// <summary>
    /// Inserts selected one-based pages from another loaded or generated PDF before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return Insert(insertBeforePageNumber, sourceDocument.GetBytesForOperation(), sourceSelection, importOptions);
    }

    /// <summary>
    /// Inserts all pages from source PDF bytes before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, byte[] sourcePdf, PdfPageImportOptions? importOptions = null) {
        return Import(sourcePdf, ImportPlacement.Insert, insertBeforePageNumber, Array.Empty<int>(), importOptions);
    }

    /// <summary>
    /// Inserts selected one-based pages from source PDF bytes before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return Import(sourcePdf, ImportPlacement.Insert, insertBeforePageNumber, GetSelectedSourcePages(sourcePdf, sourceSelection), importOptions);
    }

    /// <summary>
    /// Inserts all pages from a readable source PDF stream before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, Stream sourceStream, PdfPageImportOptions? importOptions = null) {
        return Insert(insertBeforePageNumber, ReadSourceStream(sourceStream), importOptions);
    }

    /// <summary>
    /// Inserts selected one-based pages from a readable source PDF stream before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        return Insert(insertBeforePageNumber, ReadSourceStream(sourceStream), sourceSelection, importOptions);
    }

    /// <summary>
    /// Inserts all pages from a source PDF file before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, string sourcePath, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Insert(insertBeforePageNumber, File.ReadAllBytes(sourcePath), importOptions);
    }

    /// <summary>
    /// Inserts selected one-based pages from a source PDF file before the supplied one-based page number.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public PdfDocument Insert(int insertBeforePageNumber, string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return Insert(insertBeforePageNumber, File.ReadAllBytes(sourcePath), sourceSelection, importOptions);
    }

    /// <summary>
    /// Attempts to insert pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, PdfDocument sourceDocument, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourceDocument, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert selected pages from another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, PdfDocument sourceDocument, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceDocument, nameof(sourceDocument));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourceDocument, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert source PDF bytes before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, byte[] sourcePdf, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourcePdf, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert selected pages from source PDF bytes before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, byte[] sourcePdf, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourcePdf, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert a readable source PDF stream before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, Stream sourceStream, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourceStream, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert selected pages from a readable source PDF stream before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, Stream sourceStream, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourceStream, sourceSelection, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert a source PDF file before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, string sourcePath, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourcePath, importOptions), options);
    }

    /// <summary>
    /// Attempts to insert selected pages from a source PDF file before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryInsert(int insertBeforePageNumber, string sourcePath, PdfPageSelection sourceSelection, PdfPageImportOptions? importOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        return _document.TryMutationOperation("Insert pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, () => Insert(insertBeforePageNumber, sourcePath, sourceSelection, importOptions), options);
    }

    private PdfDocument Import(byte[] sourcePdf, ImportPlacement placement, int? insertBeforePageNumber, int[] sourcePageNumbers, PdfPageImportOptions? importOptions) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        PdfPageImportOptions effectiveOptions = importOptions ?? new PdfPageImportOptions();
        byte[] targetPdf = _document.GetBytesForOperation();
        byte[] imported = placement switch {
            ImportPlacement.Append => PdfPageImporter.AppendPages(effectiveOptions, targetPdf, sourcePdf, sourcePageNumbers),
            ImportPlacement.Prepend => PdfPageImporter.PrependPages(effectiveOptions, targetPdf, sourcePdf, sourcePageNumbers),
            ImportPlacement.Insert => PdfPageImporter.InsertPages(effectiveOptions, targetPdf, sourcePdf, insertBeforePageNumber!.Value, sourcePageNumbers),
            _ => throw new ArgumentOutOfRangeException(nameof(placement), placement, "Unsupported page import placement.")
        };

        return _document.WithBytes(
            targetPdf,
            imported,
            operationName: placement.ToString());
    }

    private static int[] GetSelectedSourcePages(byte[] sourcePdf, PdfPageSelection sourceSelection) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourceSelection, nameof(sourceSelection));
        int pageCount = PdfInspector.Inspect(sourcePdf).PageCount;
        return sourceSelection.ToPageNumbers(pageCount, nameof(sourceSelection));
    }

    private static byte[] ReadSourceStream(Stream sourceStream) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        if (!sourceStream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(sourceStream));
        }

        using var buffer = new MemoryStream();
        sourceStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private enum ImportPlacement {
        Append,
        Prepend,
        Insert
    }
}
