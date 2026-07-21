namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public static byte[] ReorderPages(byte[] pdf, params int[] pageNumbers) {
        return ReorderPagesWithReadOptions(pdf, null, pageNumbers);
    }

    internal static byte[] ReorderPagesWithReadOptions(byte[] pdf, PdfReadOptions? readOptions, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        var document = PdfReadDocument.Open(pdf, readOptions);
        ValidateReorderPageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers));

        var ordered = new int[pageNumbers.Length];
        for (int i = 0; i < pageNumbers.Length; i++) {
            ordered[i] = document.Pages[pageNumbers[i] - 1].ObjectNumber;
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered, catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
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
        var bytes = ReorderPages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the specified one-based order from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, ReorderPages(File.ReadAllBytes(inputPath), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order from a file path.
    /// </summary>
    public static byte[] ReorderPages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ReorderPages(File.ReadAllBytes(inputPath), pageNumbers);
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
        var bytes = ReorderPageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with every page copied in the order produced by the specified one-based page ranges from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReorderPageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, ReorderPageRanges(File.ReadAllBytes(inputPath), pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the order produced by the specified one-based page ranges from a file path.
    /// </summary>
    public static byte[] ReorderPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ReorderPageRanges(File.ReadAllBytes(inputPath), pageRanges);
    }
}
