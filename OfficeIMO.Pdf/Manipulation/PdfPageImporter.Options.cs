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

        int targetPageCount = PdfInspector.Inspect(targetPdf, targetReadOptions).PageCount;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        PdfLoadOptions? sourceReadOptions = options.SourceReadOptions;
        byte[] preparedSource = PrepareImportSource(sourcePdf, options, sourceReadOptions);
        PdfLoadOptions? preparedSourceReadOptions = options.FlattenVisualAnnotations ? null : sourceReadOptions;
        if (insertBeforePageNumber == targetPageCount + 1) {
            return ImportPreparedPages(targetPdf, preparedSource, append: true, targetReadOptions, preparedSourceReadOptions, sourcePageNumbers);
        }

        byte[] inserted = PdfPageExtractor.ExtractPages(
            preparedSource,
            preparedSourceReadOptions,
            NormalizeSourcePageNumbers(preparedSource, sourcePageNumbers, preparedSourceReadOptions));
        if (insertBeforePageNumber == 1) {
            return PdfMerger.MergeWithPrimarySource(
                1,
                new[] { inserted, targetPdf },
                new[] { PdfLoadOptions.Default, PdfLoadOptions.Resolve(targetReadOptions) });
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber, targetReadOptions);
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
        int[] sourcePageNumbers) {
        int[] selectedPages = NormalizeSourcePageNumbers(preparedSourcePdf, sourcePageNumbers, sourceReadOptions);
        byte[] importedPages = PdfPageExtractor.ExtractPages(preparedSourcePdf, sourceReadOptions, selectedPages);
        return append
            ? PdfMerger.Merge(
                new[] { targetPdf, importedPages },
                new[] { PdfLoadOptions.Resolve(targetReadOptions), PdfLoadOptions.Default })
            : PdfMerger.MergeWithPrimarySource(
                1,
                new[] { importedPages, targetPdf },
                new[] { PdfLoadOptions.Default, PdfLoadOptions.Resolve(targetReadOptions) });
    }

    private static byte[] PrepareImportSource(byte[] sourcePdf, PdfPageImportOptions options, PdfLoadOptions? sourceReadOptions) {
        return options.FlattenVisualAnnotations
            ? PdfAnnotationFlattener.FlattenVisualAnnotations(sourcePdf, options: null, readOptions: sourceReadOptions)
            : sourcePdf;
    }
}
