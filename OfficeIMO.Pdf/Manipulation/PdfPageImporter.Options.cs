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
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        PdfReadDocument targetDocument = PdfReadDocument.Open(targetPdf, targetReadOptions);
        int targetPageCount = targetDocument.Pages.Count;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        PdfLoadOptions? sourceReadOptions = options.SourceReadOptions;
        byte[] preparedSource = PrepareImportSource(sourcePdf, options, sourceReadOptions);
        PdfLoadOptions? preparedSourceReadOptions = options.FlattenVisualAnnotations ? null : sourceReadOptions;
        if (insertBeforePageNumber == targetPageCount + 1) {
            return ImportPreparedPages(targetPdf, preparedSource, append: true, targetReadOptions, preparedSourceReadOptions, sourcePageNumbers, targetDocument);
        }

        byte[] inserted = PdfPageExtractor.ExtractPages(
            preparedSource,
            preparedSourceReadOptions,
            NormalizeSourcePageNumbers(preparedSource, sourcePageNumbers, preparedSourceReadOptions));
        if (insertBeforePageNumber == 1) {
            return MergeBoundaryPages(targetPdf, inserted, append: false, targetReadOptions, targetDocument);
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber, targetReadOptions, targetDocument);
    }

    private static byte[] ImportPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, bool append, PdfLoadOptions? targetReadOptions, int[]? sourcePageNumbers) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        PdfLoadOptions? sourceReadOptions = options.SourceReadOptions;
        byte[] preparedSource = PrepareImportSource(sourcePdf, options, sourceReadOptions);
        PdfLoadOptions? preparedSourceReadOptions = options.FlattenVisualAnnotations ? null : sourceReadOptions;
        return ImportPreparedPages(targetPdf, preparedSource, append, targetReadOptions, preparedSourceReadOptions, sourcePageNumbers!);
    }

    private static byte[] ImportPreparedPages(
        byte[] targetPdf,
        byte[] preparedSourcePdf,
        bool append,
        PdfLoadOptions? targetReadOptions,
        PdfLoadOptions? sourceReadOptions,
        int[] sourcePageNumbers,
        PdfReadDocument? targetDocument = null) {
        int[] selectedPages = NormalizeSourcePageNumbers(preparedSourcePdf, sourcePageNumbers, sourceReadOptions);
        byte[] importedPages = PdfPageExtractor.ExtractPages(preparedSourcePdf, sourceReadOptions, selectedPages);
        return MergeBoundaryPages(targetPdf, importedPages, append, targetReadOptions, targetDocument);
    }

    private static byte[] MergeBoundaryPages(
        byte[] targetPdf,
        byte[] importedPages,
        bool append,
        PdfLoadOptions? targetReadOptions,
        PdfReadDocument? targetDocument = null) {
        byte[][] sources = append ? new[] { targetPdf, importedPages } : new[] { importedPages, targetPdf };
        PdfLoadOptions[] readOptions = append
            ? new[] { PdfLoadOptions.Resolve(targetReadOptions), PdfLoadOptions.Default }
            : new[] { PdfLoadOptions.Default, PdfLoadOptions.Resolve(targetReadOptions) };
        Func<PdfReadDocument>? targetReader = targetDocument is null ? null : () => targetDocument;
        Func<PdfReadDocument>?[] readers = append
            ? new Func<PdfReadDocument>?[] { targetReader, null }
            : new Func<PdfReadDocument>?[] { null, targetReader };
        return PdfMerger.MergeWithPrimarySource(append ? 0 : 1, sources, readOptions, readers);
    }

    private static byte[] PrepareImportSource(byte[] sourcePdf, PdfPageImportOptions options, PdfLoadOptions? sourceReadOptions) {
        return options.FlattenVisualAnnotations
            ? PdfAnnotationFlattener.FlattenVisualAnnotations(sourcePdf, options: null, readOptions: sourceReadOptions)
            : sourcePdf;
    }
}
