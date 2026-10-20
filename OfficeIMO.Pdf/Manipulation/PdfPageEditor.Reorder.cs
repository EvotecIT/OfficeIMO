using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public static byte[] ReorderPages(byte[] pdf, params int[] pageNumbers) {
        return ReorderPagesWithReadOptions(pdf, null, pageNumbers);
    }

    internal static byte[] ReorderPagesWithReadOptions(byte[] pdf, PdfLoadOptions? readOptions, params int[] pageNumbers) {
        return ReorderPagesCore(pdf, readOptions, pageNumbers, null, default);
    }

    internal static byte[] ReorderPagesWithReadOptions(byte[] pdf, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        ReorderPagesCore(pdf, readOptions, pageNumbers, null, cancellationToken);

    internal static (byte[] Bytes, IReadOnlyDictionary<int, int> Map) ReorderPagesWithMapping(byte[] pdf, PdfLoadOptions? readOptions, int[] pageNumbers, CancellationToken cancellationToken = default) {
        IReadOnlyDictionary<int, int>? mapping = null;
        byte[] bytes = ReorderPagesCore(pdf, readOptions, pageNumbers, value => mapping = value, cancellationToken);
        return (bytes, mapping!);
    }

    private static byte[] ReorderPagesCore(byte[] pdf, PdfLoadOptions? readOptions, int[] pageNumbers,
        Action<IReadOnlyDictionary<int, int>>? captureObjectNumbers, CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        var (_, document) = PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.OrganizePages, readOptions, cancellationToken: cancellationToken);
        ValidateReorderPageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers));
        return PdfTaggedPageOrganizer.Rewrite(pdf, document, pageNumbers, readOptions, captureObjectNumbers: captureObjectNumbers, token: cancellationToken);
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order from the current position of a readable stream.
    /// </summary>
    public static byte[] ReorderPages(Stream stream, params int[] pageNumbers) {
        return ReorderPages(ReadStream(stream, nameof(stream)), pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the specified one-based order to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPages(byte[] pdf, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ReorderPages(pdf, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the specified one-based order from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPages(Stream inputStream, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ReorderPages(inputStream, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public static void ReorderPages(string inputPath, string outputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ReorderPages(ReadPath(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the specified one-based order from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, ReorderPages(ReadPath(inputPath), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order from a file path.
    /// </summary>
    public static byte[] ReorderPages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ReorderPages(ReadPath(inputPath), pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the order produced by the specified one-based page ranges.
    /// </summary>
    public static byte[] ReorderPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ReorderPages(pdf, ExpandPageRanges(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the order produced by the specified one-based page ranges from the current position of a readable stream.
    /// </summary>
    public static byte[] ReorderPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ReorderPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the order produced by the specified one-based page ranges to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ReorderPageRanges(pdf, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the order produced by the specified one-based page ranges from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ReorderPageRanges(inputStream, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the order produced by the specified one-based page ranges.
    /// </summary>
    public static void ReorderPageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ReorderPageRanges(ReadPath(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the order produced by the specified one-based page ranges from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, ReorderPageRanges(ReadPath(inputPath), pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the order produced by the specified one-based page ranges from a file path.
    /// </summary>
    public static byte[] ReorderPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ReorderPageRanges(ReadPath(inputPath), pageRanges);
    }
}
