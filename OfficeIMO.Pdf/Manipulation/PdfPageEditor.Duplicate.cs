namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the specified one-based pages duplicated immediately after each selected source page.
    /// Repeated selections create repeated page copies.
    /// </summary>
    public static byte[] DuplicatePages(byte[] pdf, params int[] pageNumbers) {
        return DuplicatePagesWithReadOptions(pdf, null, pageNumbers);
    }

    internal static byte[] DuplicatePagesWithReadOptions(byte[] pdf, PdfReadOptions? readOptions, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree, readOptions);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        var document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers), allowDuplicates: true);

        var duplicateCounts = new Dictionary<int, int>();
        foreach (int pageNumber in pageNumbers) {
            duplicateCounts.TryGetValue(pageNumber, out int count);
            duplicateCounts[pageNumber] = count + 1;
        }

        var ordered = new List<int>(document.Pages.Count + pageNumbers.Length);
        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            int pageObjectNumber = document.Pages[i].ObjectNumber;
            ordered.Add(pageObjectNumber);

            if (!duplicateCounts.TryGetValue(pageNumber, out int count)) {
                continue;
            }

            for (int copy = 0; copy < count; copy++) {
                ordered.Add(pageObjectNumber);
            }
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages duplicated from the current position of a readable stream.
    /// </summary>
    public static byte[] DuplicatePages(Stream stream, params int[] pageNumbers) {
        return DuplicatePages(ReadStream(stream, nameof(stream)), pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages duplicated to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePages(byte[] pdf, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, DuplicatePages(pdf, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages duplicated from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePages(Stream inputStream, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, DuplicatePages(inputStream, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages duplicated.
    /// </summary>
    public static void DuplicatePages(string inputPath, string outputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DuplicatePages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages duplicated from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DuplicatePages(File.ReadAllBytes(inputPath), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages duplicated from a file path.
    /// </summary>
    public static byte[] DuplicatePages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DuplicatePages(File.ReadAllBytes(inputPath), pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated immediately after each source page.
    /// </summary>
    public static byte[] DuplicatePageRange(byte[] pdf, int firstPage, int lastPage) {
        return DuplicatePages(pdf, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated immediately after each source page.
    /// </summary>
    public static byte[] DuplicatePageRange(byte[] pdf, PdfPageRange pageRange) {
        return DuplicatePages(pdf, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated from the current position of a readable stream.
    /// </summary>
    public static byte[] DuplicatePageRange(Stream stream, int firstPage, int lastPage) {
        return DuplicatePages(ReadStream(stream, nameof(stream)), BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated from the current position of a readable stream.
    /// </summary>
    public static byte[] DuplicatePageRange(Stream stream, PdfPageRange pageRange) {
        return DuplicatePages(ReadStream(stream, nameof(stream)), pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(byte[] pdf, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, DuplicatePageRange(pdf, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(byte[] pdf, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, DuplicatePageRange(pdf, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(Stream inputStream, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, DuplicatePageRange(inputStream, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(Stream inputStream, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, DuplicatePageRange(inputStream, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated.
    /// </summary>
    public static void DuplicatePageRange(string inputPath, string outputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DuplicatePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(string inputPath, Stream outputStream, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DuplicatePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated.
    /// </summary>
    public static void DuplicatePageRange(string inputPath, string outputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DuplicatePageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range duplicated from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void DuplicatePageRange(string inputPath, Stream outputStream, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DuplicatePageRange(File.ReadAllBytes(inputPath), pageRange));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated from a file path.
    /// </summary>
    public static byte[] DuplicatePageRange(string inputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DuplicatePageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range duplicated from a file path.
    /// </summary>
    public static byte[] DuplicatePageRange(string inputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DuplicatePageRange(File.ReadAllBytes(inputPath), pageRange);
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges duplicated immediately after each source page.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static byte[] DuplicatePageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return DuplicatePages(pdf, ExpandPageRanges(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges duplicated from the current position of a readable stream.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static byte[] DuplicatePageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return DuplicatePages(ReadStream(stream, nameof(stream)), ExpandPageRanges(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges duplicated to <paramref name="outputStream"/>.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static void DuplicatePageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, DuplicatePageRanges(pdf, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges duplicated from the current position of a readable stream to <paramref name="outputStream"/>.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static void DuplicatePageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, DuplicatePageRanges(inputStream, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges duplicated.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static void DuplicatePageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = DuplicatePageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges duplicated from a file path to <paramref name="outputStream"/>.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static void DuplicatePageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, DuplicatePageRanges(File.ReadAllBytes(inputPath), pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges duplicated from a file path.
    /// Repeated or overlapping ranges create repeated page copies.
    /// </summary>
    public static byte[] DuplicatePageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return DuplicatePageRanges(File.ReadAllBytes(inputPath), pageRanges);
    }
}
