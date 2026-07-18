namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the selected pages rotated to the specified degrees. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static byte[] RotatePages(byte[] pdf, int rotationDegrees, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree);

        int normalizedRotation = NormalizeRotation(rotationDegrees);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Open(pdf);
        var selectedPages = pageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : pageNumbers;
        ValidatePageNumbers(selectedPages, document.Pages.Count, nameof(pageNumbers));

        var selected = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (selected.Contains(pageNumber)) {
                overrides[document.Pages[i].ObjectNumber] = new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                    ["Rotate"] = new PdfNumber(normalizedRotation)
                };
            }
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF with the selected pages rotated to the specified degrees from the current position of a readable stream. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static byte[] RotatePages(Stream stream, int rotationDegrees, params int[] pageNumbers) {
        return RotatePages(ReadStream(stream, nameof(stream)), rotationDegrees, pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF with the selected pages rotated to the specified degrees to <paramref name="outputStream"/>. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static void RotatePages(byte[] pdf, Stream outputStream, int rotationDegrees, params int[] pageNumbers) {
        WriteOutput(outputStream, RotatePages(pdf, rotationDegrees, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the selected pages rotated to the specified degrees from the current position of a readable stream to <paramref name="outputStream"/>. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static void RotatePages(Stream inputStream, Stream outputStream, int rotationDegrees, params int[] pageNumbers) {
        WriteOutput(outputStream, RotatePages(inputStream, rotationDegrees, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the selected pages rotated to the specified degrees. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static void RotatePages(string inputPath, string outputPath, int rotationDegrees, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = RotatePages(File.ReadAllBytes(inputPath), rotationDegrees, pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the selected pages rotated to the specified degrees from a file path to <paramref name="outputStream"/>. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static void RotatePages(string inputPath, Stream outputStream, int rotationDegrees, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, RotatePages(File.ReadAllBytes(inputPath), rotationDegrees, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with the selected pages rotated to the specified degrees from a file path. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static byte[] RotatePages(string inputPath, int rotationDegrees, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return RotatePages(File.ReadAllBytes(inputPath), rotationDegrees, pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees.
    /// </summary>
    public static byte[] RotatePageRange(byte[] pdf, int rotationDegrees, int firstPage, int lastPage) {
        return RotatePages(pdf, rotationDegrees, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees.
    /// </summary>
    public static byte[] RotatePageRange(byte[] pdf, int rotationDegrees, PdfPageRange pageRange) {
        return RotatePages(pdf, rotationDegrees, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees from the current position of a readable stream.
    /// </summary>
    public static byte[] RotatePageRange(Stream stream, int rotationDegrees, int firstPage, int lastPage) {
        return RotatePages(ReadStream(stream, nameof(stream)), rotationDegrees, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees from the current position of a readable stream.
    /// </summary>
    public static byte[] RotatePageRange(Stream stream, int rotationDegrees, PdfPageRange pageRange) {
        return RotatePages(ReadStream(stream, nameof(stream)), rotationDegrees, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(byte[] pdf, Stream outputStream, int rotationDegrees, int firstPage, int lastPage) {
        WriteOutput(outputStream, RotatePageRange(pdf, rotationDegrees, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(byte[] pdf, Stream outputStream, int rotationDegrees, PdfPageRange pageRange) {
        WriteOutput(outputStream, RotatePageRange(pdf, rotationDegrees, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(Stream inputStream, Stream outputStream, int rotationDegrees, int firstPage, int lastPage) {
        WriteOutput(outputStream, RotatePageRange(inputStream, rotationDegrees, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(Stream inputStream, Stream outputStream, int rotationDegrees, PdfPageRange pageRange) {
        WriteOutput(outputStream, RotatePageRange(inputStream, rotationDegrees, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees.
    /// </summary>
    public static void RotatePageRange(string inputPath, string outputPath, int rotationDegrees, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(string inputPath, Stream outputStream, int rotationDegrees, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees.
    /// </summary>
    public static void RotatePageRange(string inputPath, string outputPath, int rotationDegrees, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range rotated to the specified degrees from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void RotatePageRange(string inputPath, Stream outputStream, int rotationDegrees, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees from a file path.
    /// </summary>
    public static byte[] RotatePageRange(string inputPath, int rotationDegrees, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range rotated to the specified degrees from a file path.
    /// </summary>
    public static byte[] RotatePageRange(string inputPath, int rotationDegrees, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return RotatePageRange(File.ReadAllBytes(inputPath), rotationDegrees, pageRange);
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static byte[] RotatePageRanges(byte[] pdf, int rotationDegrees, params PdfPageRange[] pageRanges) {
        return RotatePages(pdf, rotationDegrees, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees from the current position of a readable stream.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static byte[] RotatePageRanges(Stream stream, int rotationDegrees, params PdfPageRange[] pageRanges) {
        return RotatePages(ReadStream(stream, nameof(stream)), rotationDegrees, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static void RotatePageRanges(byte[] pdf, Stream outputStream, int rotationDegrees, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, RotatePageRanges(pdf, rotationDegrees, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees from the current position of a readable stream to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static void RotatePageRanges(Stream inputStream, Stream outputStream, int rotationDegrees, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, RotatePageRanges(inputStream, rotationDegrees, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static void RotatePageRanges(string inputPath, string outputPath, int rotationDegrees, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = RotatePageRanges(File.ReadAllBytes(inputPath), rotationDegrees, pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees from a file path to <paramref name="outputStream"/>.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static void RotatePageRanges(string inputPath, Stream outputStream, int rotationDegrees, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, RotatePageRanges(File.ReadAllBytes(inputPath), rotationDegrees, pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges rotated to the specified degrees from a file path.
    /// Overlapping ranges are treated as one rotation set.
    /// </summary>
    public static byte[] RotatePageRanges(string inputPath, int rotationDegrees, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return RotatePageRanges(File.ReadAllBytes(inputPath), rotationDegrees, pageRanges);
    }
}
