using OfficeIMO.Drawing.Internal;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party selected-page import helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static partial class PdfPageImporter {
    /// <summary>
    /// Appends selected one-based pages from <paramref name="sourcePdf"/> to the end of <paramref name="targetPdf"/>.
    /// When no page numbers are supplied, all source pages are appended.
    /// </summary>
    public static byte[] AppendPages(byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return ImportPages(targetPdf, sourcePdf, append: true, sourcePageNumbers);
    }

    /// <summary>
    /// Prepends selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="targetPdf"/>.
    /// When no page numbers are supplied, all source pages are prepended.
    /// </summary>
    public static byte[] PrependPages(byte[] targetPdf, byte[] sourcePdf, params int[] sourcePageNumbers) {
        return ImportPages(targetPdf, sourcePdf, append: false, sourcePageNumbers);
    }

    /// <summary>
    /// Inserts selected one-based pages from <paramref name="sourcePdf"/> before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>.
    /// Use target page count + 1 to insert at the end. When no page numbers are supplied, all source pages are inserted.
    /// </summary>
    public static byte[] InsertPages(byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        int targetPageCount = PdfInspector.Inspect(targetPdf).PageCount;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        if (insertBeforePageNumber == targetPageCount + 1) {
            return AppendPages(targetPdf, sourcePdf, sourcePageNumbers);
        }

        byte[] inserted = PdfPageExtractor.ExtractPages(sourcePdf, NormalizeSourcePageNumbers(sourcePdf, sourcePageNumbers));
        if (insertBeforePageNumber == 1) {
            return PdfMerger.MergeWithPrimarySource(1, inserted, targetPdf);
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber);
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        return InsertPages(targetPdf, sourcePdf, insertBeforePageNumber, BuildInclusivePageRange(firstSourcePage, lastSourcePage, nameof(lastSourcePage)));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        return InsertPages(targetPdf, sourcePdf, insertBeforePageNumber, sourcePageRange.ToPageNumbers());
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges to the end of <paramref name="targetPdf"/>.
    /// Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] AppendPageRanges(byte[] targetPdf, byte[] sourcePdf, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        byte[] importedPages = PdfPageExtractor.ExtractPageRanges(sourcePdf, sourcePageRanges);
        return PdfMerger.Merge(targetPdf, importedPages);
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges before <paramref name="targetPdf"/>.
    /// Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] PrependPageRanges(byte[] targetPdf, byte[] sourcePdf, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        byte[] importedPages = PdfPageExtractor.ExtractPageRanges(sourcePdf, sourcePageRanges);
        return PdfMerger.MergeWithPrimarySource(1, importedPages, targetPdf);
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges before <paramref name="insertBeforePageNumber"/> in <paramref name="targetPdf"/>.
    /// Use target page count + 1 to insert at the end. Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] InsertPageRanges(byte[] targetPdf, byte[] sourcePdf, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        int targetPageCount = PdfInspector.Inspect(targetPdf).PageCount;
        ValidateInsertBeforePageNumber(insertBeforePageNumber, targetPageCount);

        if (insertBeforePageNumber == targetPageCount + 1) {
            return AppendPageRanges(targetPdf, sourcePdf, sourcePageRanges);
        }

        byte[] inserted = PdfPageExtractor.ExtractPageRanges(sourcePdf, sourcePageRanges);
        if (insertBeforePageNumber == 1) {
            return PdfMerger.MergeWithPrimarySource(1, inserted, targetPdf);
        }

        return PdfMerger.MergePrimaryWithInsertedPages(targetPdf, inserted, insertBeforePageNumber);
    }

    /// <summary>
    /// Appends selected one-based pages from a readable source stream to the end of a readable target stream.
    /// Streams are read from their current positions. When no page numbers are supplied, all source pages are appended.
    /// </summary>
    public static byte[] AppendPages(Stream targetStream, Stream sourceStream, params int[] sourcePageNumbers) {
        return AppendPages(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), sourcePageNumbers);
    }

    /// <summary>
    /// Prepends selected one-based pages from a readable source stream before a readable target stream.
    /// Streams are read from their current positions. When no page numbers are supplied, all source pages are prepended.
    /// </summary>
    public static byte[] PrependPages(Stream targetStream, Stream sourceStream, params int[] sourcePageNumbers) {
        return PrependPages(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), sourcePageNumbers);
    }

    /// <summary>
    /// Inserts selected one-based pages from a readable source stream before <paramref name="insertBeforePageNumber"/> in a readable target stream.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPages(Stream targetStream, Stream sourceStream, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        return InsertPages(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), insertBeforePageNumber, sourcePageNumbers);
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a readable target stream.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(Stream targetStream, Stream sourceStream, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        return InsertPageRange(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), insertBeforePageNumber, firstSourcePage, lastSourcePage);
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a readable target stream.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(Stream targetStream, Stream sourceStream, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        return InsertPageRange(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), insertBeforePageNumber, sourcePageRange);
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges from a readable source stream to the end of a readable target stream.
    /// Streams are read from their current positions. Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] AppendPageRanges(Stream targetStream, Stream sourceStream, params PdfPageRange[] sourcePageRanges) {
        return AppendPageRanges(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), sourcePageRanges);
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges from a readable source stream before a readable target stream.
    /// Streams are read from their current positions. Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] PrependPageRanges(Stream targetStream, Stream sourceStream, params PdfPageRange[] sourcePageRanges) {
        return PrependPageRanges(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), sourcePageRanges);
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges from a readable source stream before <paramref name="insertBeforePageNumber"/> in a readable target stream.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRanges(Stream targetStream, Stream sourceStream, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        return InsertPageRanges(ReadStream(targetStream, nameof(targetStream)), ReadStream(sourceStream, nameof(sourceStream)), insertBeforePageNumber, sourcePageRanges);
    }

    /// <summary>
    /// Appends selected one-based pages from <paramref name="sourcePdf"/> and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void AppendPages(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, AppendPages(targetPdf, sourcePdf, sourcePageNumbers));
    }

    /// <summary>
    /// Prepends selected one-based pages from <paramref name="sourcePdf"/> and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void PrependPages(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, PrependPages(targetPdf, sourcePdf, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts selected one-based pages from <paramref name="sourcePdf"/> and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPages(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, InsertPages(targetPdf, sourcePdf, insertBeforePageNumber, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRange(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        WriteOutput(outputStream, InsertPageRange(targetPdf, sourcePdf, insertBeforePageNumber, firstSourcePage, lastSourcePage));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRange(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        WriteOutput(outputStream, InsertPageRange(targetPdf, sourcePdf, insertBeforePageNumber, sourcePageRange));
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void AppendPageRanges(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, AppendPageRanges(targetPdf, sourcePdf, sourcePageRanges));
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void PrependPageRanges(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, PrependPageRanges(targetPdf, sourcePdf, sourcePageRanges));
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRanges(byte[] targetPdf, byte[] sourcePdf, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, InsertPageRanges(targetPdf, sourcePdf, insertBeforePageNumber, sourcePageRanges));
    }

    /// <summary>
    /// Appends selected one-based pages from a readable source stream to a readable target stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions.
    /// </summary>
    public static void AppendPages(Stream targetStream, Stream sourceStream, Stream outputStream, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, AppendPages(targetStream, sourceStream, sourcePageNumbers));
    }

    /// <summary>
    /// Prepends selected one-based pages from a readable source stream before a readable target stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions.
    /// </summary>
    public static void PrependPages(Stream targetStream, Stream sourceStream, Stream outputStream, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, PrependPages(targetStream, sourceStream, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts selected one-based pages from a readable source stream before <paramref name="insertBeforePageNumber"/> in a readable target stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static void InsertPages(Stream targetStream, Stream sourceStream, Stream outputStream, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        WriteOutput(outputStream, InsertPages(targetStream, sourceStream, insertBeforePageNumber, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range from a readable source stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static void InsertPageRange(Stream targetStream, Stream sourceStream, Stream outputStream, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        WriteOutput(outputStream, InsertPageRange(targetStream, sourceStream, insertBeforePageNumber, firstSourcePage, lastSourcePage));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range from a readable source stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static void InsertPageRange(Stream targetStream, Stream sourceStream, Stream outputStream, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        WriteOutput(outputStream, InsertPageRange(targetStream, sourceStream, insertBeforePageNumber, sourcePageRange));
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges from a readable source stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions.
    /// </summary>
    public static void AppendPageRanges(Stream targetStream, Stream sourceStream, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, AppendPageRanges(targetStream, sourceStream, sourcePageRanges));
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges from a readable source stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions.
    /// </summary>
    public static void PrependPageRanges(Stream targetStream, Stream sourceStream, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, PrependPageRanges(targetStream, sourceStream, sourcePageRanges));
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges from a readable source stream and writes the result to <paramref name="outputStream"/>.
    /// Streams are read from their current positions. Use target page count + 1 to insert at the end.
    /// </summary>
    public static void InsertPageRanges(Stream targetStream, Stream sourceStream, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        WriteOutput(outputStream, InsertPageRanges(targetStream, sourceStream, insertBeforePageNumber, sourcePageRanges));
    }

    /// <summary>
    /// Appends selected one-based pages from a source PDF file to a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void AppendPages(string targetPath, string sourcePath, Stream outputStream, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, AppendPages(targetPath, sourcePath, sourcePageNumbers));
    }

    /// <summary>
    /// Prepends selected one-based pages from a source PDF file before a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void PrependPages(string targetPath, string sourcePath, Stream outputStream, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, PrependPages(targetPath, sourcePath, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts selected one-based pages from a source PDF file before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPages(string targetPath, string sourcePath, Stream outputStream, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, InsertPages(targetPath, sourcePath, insertBeforePageNumber, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRange(string targetPath, string sourcePath, Stream outputStream, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, InsertPageRange(targetPath, sourcePath, insertBeforePageNumber, firstSourcePage, lastSourcePage));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRange(string targetPath, string sourcePath, Stream outputStream, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, InsertPageRange(targetPath, sourcePath, insertBeforePageNumber, sourcePageRange));
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges from a source PDF file to a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void AppendPageRanges(string targetPath, string sourcePath, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, AppendPageRanges(targetPath, sourcePath, sourcePageRanges));
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges from a source PDF file before a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void PrependPageRanges(string targetPath, string sourcePath, Stream outputStream, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, PrependPageRanges(targetPath, sourcePath, sourcePageRanges));
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges from a source PDF file before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void InsertPageRanges(string targetPath, string sourcePath, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, InsertPageRanges(targetPath, sourcePath, insertBeforePageNumber, sourcePageRanges));
    }

    /// <summary>
    /// Appends selected one-based pages from a source PDF file to a target PDF file and returns the imported PDF bytes.
    /// When no page numbers are supplied, all source pages are appended.
    /// </summary>
    public static byte[] AppendPages(string targetPath, string sourcePath, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        return AppendPages(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), sourcePageNumbers);
    }

    /// <summary>
    /// Prepends selected one-based pages from a source PDF file before a target PDF file and returns the imported PDF bytes.
    /// When no page numbers are supplied, all source pages are prepended.
    /// </summary>
    public static byte[] PrependPages(string targetPath, string sourcePath, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        return PrependPages(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), sourcePageNumbers);
    }

    /// <summary>
    /// Inserts selected one-based pages from a source PDF file before <paramref name="insertBeforePageNumber"/> in a target PDF file and returns the imported PDF bytes.
    /// Use target page count + 1 to insert at the end. When no page numbers are supplied, all source pages are inserted.
    /// </summary>
    public static byte[] InsertPages(string targetPath, string sourcePath, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        return InsertPages(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), insertBeforePageNumber, sourcePageNumbers);
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and returns the imported PDF bytes.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(string targetPath, string sourcePath, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        return InsertPageRange(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), insertBeforePageNumber, firstSourcePage, lastSourcePage);
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and returns the imported PDF bytes.
    /// Use target page count + 1 to insert at the end.
    /// </summary>
    public static byte[] InsertPageRange(string targetPath, string sourcePath, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        return InsertPageRange(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), insertBeforePageNumber, sourcePageRange);
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges from a source PDF file to a target PDF file and returns the imported PDF bytes.
    /// Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] AppendPageRanges(string targetPath, string sourcePath, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        return AppendPageRanges(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), sourcePageRanges);
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges from a source PDF file before a target PDF file and returns the imported PDF bytes.
    /// Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] PrependPageRanges(string targetPath, string sourcePath, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        return PrependPageRanges(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), sourcePageRanges);
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges before <paramref name="insertBeforePageNumber"/> in a target PDF file and returns the imported PDF bytes.
    /// Use target page count + 1 to insert at the end. Ranges are imported in caller order; repeated or overlapping ranges create cloned source pages.
    /// </summary>
    public static byte[] InsertPageRanges(string targetPath, string sourcePath, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        Guard.NotNull(sourcePageRanges, nameof(sourcePageRanges));

        return InsertPageRanges(File.ReadAllBytes(targetPath), File.ReadAllBytes(sourcePath), insertBeforePageNumber, sourcePageRanges);
    }

    /// <summary>
    /// Appends selected one-based pages from a source PDF file to a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void AppendPages(string targetPath, string sourcePath, string outputPath, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), AppendPages(targetPath, sourcePath, sourcePageNumbers));
    }

    /// <summary>
    /// Prepends selected one-based pages from a source PDF file before a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void PrependPages(string targetPath, string sourcePath, string outputPath, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), PrependPages(targetPath, sourcePath, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts selected one-based pages from a source PDF file before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void InsertPages(string targetPath, string sourcePath, string outputPath, int insertBeforePageNumber, params int[] sourcePageNumbers) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), InsertPages(targetPath, sourcePath, insertBeforePageNumber, sourcePageNumbers));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void InsertPageRange(string targetPath, string sourcePath, string outputPath, int insertBeforePageNumber, int firstSourcePage, int lastSourcePage) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), InsertPageRange(targetPath, sourcePath, insertBeforePageNumber, firstSourcePage, lastSourcePage));
    }

    /// <summary>
    /// Inserts the inclusive one-based source page range before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void InsertPageRange(string targetPath, string sourcePath, string outputPath, int insertBeforePageNumber, PdfPageRange sourcePageRange) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), InsertPageRange(targetPath, sourcePath, insertBeforePageNumber, sourcePageRange));
    }

    /// <summary>
    /// Appends the supplied inclusive one-based source page ranges from a source PDF file to a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void AppendPageRanges(string targetPath, string sourcePath, string outputPath, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), AppendPageRanges(targetPath, sourcePath, sourcePageRanges));
    }

    /// <summary>
    /// Prepends the supplied inclusive one-based source page ranges from a source PDF file before a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void PrependPageRanges(string targetPath, string sourcePath, string outputPath, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), PrependPageRanges(targetPath, sourcePath, sourcePageRanges));
    }

    /// <summary>
    /// Inserts the supplied inclusive one-based source page ranges from a source PDF file before <paramref name="insertBeforePageNumber"/> in a target PDF file and writes the result to <paramref name="outputPath"/>.
    /// </summary>
    public static void InsertPageRanges(string targetPath, string sourcePath, string outputPath, int insertBeforePageNumber, params PdfPageRange[] sourcePageRanges) {
        Guard.NotNullOrWhiteSpace(targetPath, nameof(targetPath));
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));

        WriteOutput(ValidateOutputPath(outputPath), InsertPageRanges(targetPath, sourcePath, insertBeforePageNumber, sourcePageRanges));
    }

    private static byte[] ImportPages(byte[] targetPdf, byte[] sourcePdf, bool append, int[]? sourcePageNumbers) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(sourcePageNumbers, nameof(sourcePageNumbers));

        int[] selectedPages = NormalizeSourcePageNumbers(sourcePdf, sourcePageNumbers!);
        byte[] importedPages = PdfPageExtractor.ExtractPages(sourcePdf, selectedPages);
        return append
            ? PdfMerger.Merge(targetPdf, importedPages)
            : PdfMerger.MergeWithPrimarySource(1, importedPages, targetPdf);
    }

    private static int[] NormalizeSourcePageNumbers(byte[] sourcePdf, int[] sourcePageNumbers, PdfReadOptions? sourceReadOptions = null) {
        if (sourcePageNumbers.Length > 0) {
            return sourcePageNumbers;
        }

        PdfDocumentInfo info = PdfInspector.Inspect(sourcePdf, sourceReadOptions);
        if (info.PageCount == 0) {
            throw new ArgumentException("Source PDF does not contain any pages.", nameof(sourcePdf));
        }

        return Enumerable.Range(1, info.PageCount).ToArray();
    }

    private static void ValidateInsertBeforePageNumber(int insertBeforePageNumber, int pageCount) {
        if (insertBeforePageNumber < 1 || insertBeforePageNumber > pageCount + 1) {
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber), "Insert-before page must be in the target document page range 1-" + (pageCount + 1).ToString(CultureInfo.InvariantCulture) + ".");
        }
    }

    private static int[] BuildInclusivePageRange(int firstPage, int lastPage, string lastPageParamName) {
        if (firstPage > lastPage) {
            throw new ArgumentOutOfRangeException(lastPageParamName, "Last page must be greater than or equal to first page.");
        }

        return Enumerable.Range(firstPage, lastPage - firstPage + 1).ToArray();
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        ValidateWritableOutputStream(outputStream);
        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }
}
