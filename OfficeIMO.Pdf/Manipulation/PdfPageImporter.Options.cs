using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImporter {
    /// <summary>
    /// Appends selected one-based pages from <paramref name="sourcePdf"/> to the end of <paramref name="targetPdf"/>, applying optional source preparation first.
    /// When no page numbers are supplied, all source pages are appended.
    /// </summary>
    public static byte[] AppendPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return AppendPages(options, targetPdf, sourcePdf, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] AppendPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        return ImportPages(options, targetPdf, sourcePdf, append: true, targetReadOptions, sourcePageNumbers);
    }

    /// <summary>
    /// Prepends selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="targetPdf"/>, applying optional source preparation first.
    /// When no page numbers are supplied, all source pages are prepended.
    /// </summary>
    public static byte[] PrependPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return PrependPages(options, targetPdf, sourcePdf, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] PrependPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        return ImportPages(options, targetPdf, sourcePdf, append: false, targetReadOptions, sourcePageNumbers);
    }

    /// <summary>
    /// Inserts selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>, applying optional source preparation first.
    /// Use target page count + 1 to insert at the end. When no page numbers are supplied, all source pages are inserted.
    /// </summary>
    public static byte[] InsertPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        return InsertPages(options, targetPdf, sourcePdf, insertBeforePageNumber, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] InsertPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        Guard.NotNull(options, nameof(options));
        CancellationToken cancellationToken = options.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        PdfReadDocument targetDocument = PdfReadDocument.Open(targetPdf, targetReadOptions, cancellationToken);
        int targetPageCount = targetDocument.Pages.Count;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        PdfLoadOptions? sourceReadOptions = options.SourceReadOptions;
        byte[] preparedSource = PrepareImportSource(sourcePdf, options, sourceReadOptions);
        PdfLoadOptions preparedSourceReadOptions = PdfLoadOptions.WithMinimumInputBytes(sourceReadOptions, preparedSource.LongLength);
        if (TryImportPreservingCatalog(targetPdf, preparedSource, insertBeforePageNumber, sourcePageNumbers, targetReadOptions, preparedSourceReadOptions, out var taggedOutput, cancellationToken))
            return taggedOutput;
        if (insertBeforePageNumber == targetPageCount + 1) {
            return ImportPreparedPages(targetPdf, preparedSource, append: true, targetReadOptions, preparedSourceReadOptions, sourcePageNumbers, targetDocument, cancellationToken);
        }

        byte[] inserted = ExtractImportedPages(preparedSource, preparedSourceReadOptions, sourcePageNumbers, cancellationToken);
        if (insertBeforePageNumber == 1) {
            return MergeBoundaryPages(targetPdf, inserted, append: false, targetReadOptions, targetDocument, preparedSourceReadOptions, cancellationToken);
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber, targetReadOptions, targetDocument, preparedSourceReadOptions, cancellationToken);
    }

    private static byte[] ImportPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, bool append, PdfLoadOptions? targetReadOptions, int[]? sourcePageNumbers) {
        Guard.NotNull(options, nameof(options));
        CancellationToken cancellationToken = options.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        PdfLoadOptions? sourceReadOptions = options.SourceReadOptions;
        byte[] preparedSource = PrepareImportSource(sourcePdf, options, sourceReadOptions);
        PdfLoadOptions preparedSourceReadOptions = PdfLoadOptions.WithMinimumInputBytes(sourceReadOptions, preparedSource.LongLength);
        var target = PdfReadDocument.Open(targetPdf, targetReadOptions, cancellationToken);
        if (TryImportPreservingCatalog(targetPdf, preparedSource, append ? target.Pages.Count + 1 : 1, sourcePageNumbers!, targetReadOptions, preparedSourceReadOptions, out var taggedOutput, cancellationToken))
            return taggedOutput;
        return ImportPreparedPages(targetPdf, preparedSource, append, targetReadOptions, preparedSourceReadOptions, sourcePageNumbers!, cancellationToken: cancellationToken);
    }

    private static bool TryImportPreservingCatalog(byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, int[] sourcePageNumbers,
        PdfLoadOptions? targetOptions, PdfLoadOptions? sourceOptions, out byte[] output, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var target = PdfReadDocument.Open(targetPdf, targetOptions, cancellationToken);
        var source = PdfReadDocument.Open(sourcePdf, sourceOptions, cancellationToken);
        if (!PdfTaggedPageOrganizer.IsTagged(target) && !PdfTaggedPageOrganizer.IsTagged(source)) {
            var sourceCatalog = (PdfDictionary)source.Objects[source.Security.RootObjectNumber!.Value].Value;
            // Incoming forms and layers still need the merger's explicit collision policies. Ordinary page graphs
            // use the same catalog-preserving ImportPages route whether either document is tagged or not.
            bool needsMergePolicy = source.AcroFormXfa != null || source.FormFields.Count > 0 || sourceCatalog.Items.ContainsKey("OCProperties") ||
                sourceCatalog.Items.TryGetValue("Names", out var value) && PdfObjectLookup.Resolve(source.Objects, value) is PdfDictionary names &&
                names.Items.Keys.Any(key => key != "Dests");
            if (needsMergePolicy) { output = Array.Empty<byte>(); return false; }
        }
        output = PdfTaggedPageOrganizer.Import(targetPdf, sourcePdf, insertBeforePageNumber, sourcePageNumbers, targetOptions, sourceOptions, cancellationToken);
        return true;
    }

    private static byte[] ImportPreparedPages(
        byte[] targetPdf,
        byte[] preparedSourcePdf,
        bool append,
        PdfLoadOptions? targetReadOptions,
        PdfLoadOptions? sourceReadOptions,
        int[] sourcePageNumbers,
        PdfReadDocument? targetDocument = null, CancellationToken cancellationToken = default) {
        byte[] importedPages = ExtractImportedPages(preparedSourcePdf, sourceReadOptions, sourcePageNumbers, cancellationToken);
        return MergeBoundaryPages(targetPdf, importedPages, append, targetReadOptions, targetDocument, sourceReadOptions, cancellationToken);
    }

    private static byte[] ExtractImportedPages(byte[] sourcePdf, PdfLoadOptions? sourceReadOptions, int[] sourcePageNumbers, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (sourcePageNumbers.Length != 0) {
            return PdfPageExtractor.ExtractPages(sourcePdf, sourcePageNumbers, sourceReadOptions, documentFactory: null, cancellationToken);
        }

        PdfReadDocument sourceDocument = PdfReadDocument.Open(sourcePdf, sourceReadOptions, cancellationToken);
        if (sourceDocument.Pages.Count == 0) {
            throw new ArgumentException("Source PDF does not contain any pages.", nameof(sourcePdf));
        }

        return PdfPageExtractor.ExtractPages(
            sourcePdf,
            Enumerable.Range(1, sourceDocument.Pages.Count),
            sourceReadOptions,
            () => sourceDocument, cancellationToken);
    }

    private static byte[] MergeBoundaryPages(
        byte[] targetPdf,
        byte[] importedPages,
        bool append,
        PdfLoadOptions? targetReadOptions,
        PdfReadDocument? targetDocument = null,
        PdfLoadOptions? importedReadOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[][] sources = append ? new[] { targetPdf, importedPages } : new[] { importedPages, targetPdf };
        PdfLoadOptions[] readOptions = append
            ? new[] { PdfLoadOptions.Resolve(targetReadOptions), PdfLoadOptions.WithMinimumInputBytes(importedReadOptions, importedPages.LongLength) }
            : new[] { PdfLoadOptions.WithMinimumInputBytes(importedReadOptions, importedPages.LongLength), PdfLoadOptions.Resolve(targetReadOptions) };
        Func<PdfReadDocument>? targetReader = targetDocument is null ? null : () => targetDocument;
        Func<PdfReadDocument>?[] readers = append
            ? new Func<PdfReadDocument>?[] { targetReader, null }
            : new Func<PdfReadDocument>?[] { null, targetReader };
        return PdfMerger.MergeWithPrimarySource(append ? 0 : 1, sources, readOptions, readers, cancellationToken);
    }

    private static byte[] PrepareImportSource(byte[] sourcePdf, PdfPageImportOptions options, PdfLoadOptions? sourceReadOptions) {
        options.CancellationToken.ThrowIfCancellationRequested();
        _ = PdfDocumentSource.FromOwnedBytes(sourcePdf, sourceReadOptions);
        return options.FlattenVisualAnnotations
            ? PdfAnnotationFlattener.FlattenVisualAnnotations(sourcePdf, new PdfAnnotationFlattenOptions { CancellationToken = options.CancellationToken }, readOptions: sourceReadOptions)
            : sourcePdf;
    }
}
