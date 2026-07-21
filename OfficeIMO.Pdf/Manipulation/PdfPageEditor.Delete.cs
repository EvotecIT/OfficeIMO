namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the specified one-based pages removed.
    /// </summary>
    public static byte[] DeletePages(byte[] pdf, params int[] pageNumbers) {
        return DeletePagesWithReadOptions(pdf, readOptions: null, pageNumbers);
    }

    internal static byte[] DeletePagesWithReadOptions(byte[] pdf, PdfReadOptions? readOptions, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree, readOptions);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        var document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers));

        var deleted = new HashSet<int>(pageNumbers);
        if (deleted.Count == document.Pages.Count) {
            throw new ArgumentException("Cannot delete every page from a PDF.", nameof(pageNumbers));
        }

        var remaining = new List<int>(document.Pages.Count - deleted.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!deleted.Contains(pageNumber)) {
                remaining.Add(document.Pages[i].ObjectNumber);
            }
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.UncheckedMetadata, remaining.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages removed from the current position of a readable stream.
    /// </summary>
    public static byte[] DeletePages(Stream stream, params int[] pageNumbers) {
        return DeletePages(ReadStream(stream, nameof(stream)), pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages removed to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePages(byte[] pdf, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, DeletePages(pdf, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages removed from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePages(Stream inputStream, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, DeletePages(inputStream, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages removed.
    /// </summary>
    public static void DeletePages(string inputPath, string outputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DeletePages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages removed from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DeletePages(File.ReadAllBytes(inputPath), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages removed from a file path.
    /// </summary>
    public static byte[] DeletePages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DeletePages(File.ReadAllBytes(inputPath), pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed.
    /// </summary>
    public static byte[] DeletePageRange(byte[] pdf, int firstPage, int lastPage) {
        return DeletePages(pdf, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed.
    /// </summary>
    public static byte[] DeletePageRange(byte[] pdf, PdfPageRange pageRange) {
        return DeletePages(pdf, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed from the current position of a readable stream.
    /// </summary>
    public static byte[] DeletePageRange(Stream stream, int firstPage, int lastPage) {
        return DeletePages(ReadStream(stream, nameof(stream)), BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed from the current position of a readable stream.
    /// </summary>
    public static byte[] DeletePageRange(Stream stream, PdfPageRange pageRange) {
        return DeletePages(ReadStream(stream, nameof(stream)), pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(byte[] pdf, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, DeletePageRange(pdf, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(byte[] pdf, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, DeletePageRange(pdf, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(Stream inputStream, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, DeletePageRange(inputStream, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(Stream inputStream, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, DeletePageRange(inputStream, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed.
    /// </summary>
    public static void DeletePageRange(string inputPath, string outputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DeletePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(string inputPath, Stream outputStream, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DeletePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed.
    /// </summary>
    public static void DeletePageRange(string inputPath, string outputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DeletePageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range removed from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DeletePageRange(string inputPath, Stream outputStream, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DeletePageRange(File.ReadAllBytes(inputPath), pageRange));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed from a file path.
    /// </summary>
    public static byte[] DeletePageRange(string inputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DeletePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range removed from a file path.
    /// </summary>
    public static byte[] DeletePageRange(string inputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DeletePageRange(File.ReadAllBytes(inputPath), pageRange);
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges removed.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static byte[] DeletePageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return DeletePages(pdf, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges removed from the current position of a readable stream.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static byte[] DeletePageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return DeletePages(ReadStream(stream, nameof(stream)), ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges removed to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static void DeletePageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, DeletePageRanges(pdf, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges removed from the current position of a readable stream to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static void DeletePageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, DeletePageRanges(inputStream, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges removed.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static void DeletePageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DeletePageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges removed from a file path to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static void DeletePageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DeletePageRanges(File.ReadAllBytes(inputPath), pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges removed from a file path.
    /// Overlapping ranges are treated as one deletion set.
    /// </summary>
    public static byte[] DeletePageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DeletePageRanges(File.ReadAllBytes(inputPath), pageRanges);
    }
}
