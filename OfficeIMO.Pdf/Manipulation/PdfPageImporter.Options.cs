namespace OfficeIMO.Pdf;

internal static partial class PdfPageImporter {
    /// <summary>
    /// Appends selected one-based pages from <paramref name="sourcePdf"/> to the end of <paramref name="targetPdf"/>, applying optional source preparation first.
    /// When no page numbers are supplied, all source pages are appended.
    /// </summary>
    public static byte[] AppendPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return AppendPages(options, targetPdf, sourcePdf, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] AppendPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, PdfReadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        return ImportPages(options, targetPdf, sourcePdf, append: true, targetReadOptions, sourcePageNumbers);
    }

    /// <summary>
    /// Prepends selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="targetPdf"/>, applying optional source preparation first.
    /// When no page numbers are supplied, all source pages are prepended.
    /// </summary>
    public static byte[] PrependPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return PrependPages(options, targetPdf, sourcePdf, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] PrependPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, PdfReadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        return ImportPages(options, targetPdf, sourcePdf, append: false, targetReadOptions, sourcePageNumbers);
    }

    /// <summary>
    /// Inserts selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>, applying optional source preparation first.
    /// Use target page count + 1 to insert at the end. When no page numbers are supplied, all source pages are inserted.
    /// </summary>
    public static byte[] InsertPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        return InsertPages(options, targetPdf, sourcePdf, insertBeforePageNumber, targetReadOptions: null, sourcePageNumbers);
    }

    internal static byte[] InsertPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, PdfReadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        int targetPageCount = PdfInspector.Inspect(targetPdf, targetReadOptions).PageCount;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        byte[] preparedSource = PrepareImportSource(sourcePdf, options);
        if (insertBeforePageNumber == targetPageCount + 1) {
            return ImportPreparedPages(targetPdf, preparedSource, append: true, targetReadOptions, sourcePageNumbers);
        }

        byte[] inserted = PdfPageExtractor.ExtractPages(preparedSource, NormalizeSourcePageNumbers(preparedSource, sourcePageNumbers));
        if (insertBeforePageNumber == 1) {
            return PdfMerger.MergeWithPrimarySource(
                1,
                new[] { inserted, targetPdf },
                new[] { PdfReadOptions.Default, PdfReadOptions.Resolve(targetReadOptions) });
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber, targetReadOptions);
    }

    private static byte[] ImportPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, bool append, PdfReadOptions? targetReadOptions, int[]? sourcePageNumbers) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        return ImportPreparedPages(targetPdf, PrepareImportSource(sourcePdf, options), append, targetReadOptions, sourcePageNumbers!);
    }

    private static byte[] ImportPreparedPages(byte[] targetPdf, byte[] preparedSourcePdf, bool append, PdfReadOptions? targetReadOptions, int[] sourcePageNumbers) {
        int[] selectedPages = NormalizeSourcePageNumbers(preparedSourcePdf, sourcePageNumbers);
        byte[] importedPages = PdfPageExtractor.ExtractPages(preparedSourcePdf, selectedPages);
        return append
            ? PdfMerger.Merge(
                new[] { targetPdf, importedPages },
                new[] { PdfReadOptions.Resolve(targetReadOptions), PdfReadOptions.Default })
            : PdfMerger.MergeWithPrimarySource(
                1,
                new[] { importedPages, targetPdf },
                new[] { PdfReadOptions.Default, PdfReadOptions.Resolve(targetReadOptions) });
    }

    private static byte[] PrepareImportSource(byte[] sourcePdf, PdfPageImportOptions options) {
        return options.FlattenVisualAnnotations
            ? PdfAnnotationFlattener.FlattenVisualAnnotations(sourcePdf)
            : sourcePdf;
    }
}
