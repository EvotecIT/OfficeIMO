namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the specified one-based pages moved before <paramref name="insertBeforePageNumber"/>.
    /// The moved pages keep their original relative order. Use page count + 1 to move pages to the end.
    /// </summary>
    public static byte[] MovePages(byte[] pdf, int insertBeforePageNumber, params int[] pageNumbers) {
        return MovePagesWithReadOptions(pdf, insertBeforePageNumber, readOptions: null, pageNumbers);
    }

    internal static byte[] MovePagesWithReadOptions(byte[] pdf, int insertBeforePageNumber, PdfReadOptions? readOptions, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree, readOptions);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        var document = PdfReadDocument.Open(pdf, readOptions);
        ValidateMoveInsertBeforePageNumber(insertBeforePageNumber, document.Pages.Count);
        ValidatePageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers));

        var selected = new HashSet<int>(pageNumbers);
        if (insertBeforePageNumber <= document.Pages.Count && selected.Contains(insertBeforePageNumber)) {
            throw new ArgumentException("Insert-before page cannot be one of the moved pages.", nameof(insertBeforePageNumber));
        }

        var moving = new List<int>(selected.Count);
        var remaining = new List<(int PageNumber, int PageObjectNumber)>(document.Pages.Count - selected.Count);

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            int pageObjectNumber = document.Pages[i].ObjectNumber;
            if (selected.Contains(pageNumber)) {
                moving.Add(pageObjectNumber);
            } else {
                remaining.Add((pageNumber, pageObjectNumber));
            }
        }

        int insertionIndex = insertBeforePageNumber == document.Pages.Count + 1
            ? remaining.Count
            : remaining.TakeWhile(page => page.PageNumber < insertBeforePageNumber).Count();

        var ordered = new List<int>(document.Pages.Count);
        for (int i = 0; i < insertionIndex; i++) {
            ordered.Add(remaining[i].PageObjectNumber);
        }

        ordered.AddRange(moving);

        for (int i = insertionIndex; i < remaining.Count; i++) {
            ordered.Add(remaining[i].PageObjectNumber);
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages moved from the current position of a readable stream.
    /// </summary>
    public static byte[] MovePages(Stream stream, int insertBeforePageNumber, params int[] pageNumbers) {
        return MovePages(ReadStream(stream, nameof(stream)), insertBeforePageNumber, pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages moved to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePages(byte[] pdf, Stream outputStream, int insertBeforePageNumber, params int[] pageNumbers) {
        WriteOutput(outputStream, MovePages(pdf, insertBeforePageNumber, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages moved from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePages(Stream inputStream, Stream outputStream, int insertBeforePageNumber, params int[] pageNumbers) {
        WriteOutput(outputStream, MovePages(inputStream, insertBeforePageNumber, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages moved.
    /// </summary>
    public static void MovePages(string inputPath, string outputPath, int insertBeforePageNumber, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = MovePages(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the specified one-based pages moved from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePages(string inputPath, Stream outputStream, int insertBeforePageNumber, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, MovePages(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with the specified one-based pages moved from a file path.
    /// </summary>
    public static byte[] MovePages(string inputPath, int insertBeforePageNumber, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return MovePages(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved before <paramref name="insertBeforePageNumber"/>.
    /// The moved pages keep their original relative order. Use page count + 1 to move pages to the end.
    /// </summary>
    public static byte[] MovePageRange(byte[] pdf, int insertBeforePageNumber, int firstPage, int lastPage) {
        return MovePages(pdf, insertBeforePageNumber, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved before <paramref name="insertBeforePageNumber"/>.
    /// The moved pages keep their original relative order. Use page count + 1 to move pages to the end.
    /// </summary>
    public static byte[] MovePageRange(byte[] pdf, int insertBeforePageNumber, PdfPageRange pageRange) {
        return MovePages(pdf, insertBeforePageNumber, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved from the current position of a readable stream.
    /// </summary>
    public static byte[] MovePageRange(Stream stream, int insertBeforePageNumber, int firstPage, int lastPage) {
        return MovePages(ReadStream(stream, nameof(stream)), insertBeforePageNumber, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved from the current position of a readable stream.
    /// </summary>
    public static byte[] MovePageRange(Stream stream, int insertBeforePageNumber, PdfPageRange pageRange) {
        return MovePages(ReadStream(stream, nameof(stream)), insertBeforePageNumber, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(byte[] pdf, Stream outputStream, int insertBeforePageNumber, int firstPage, int lastPage) {
        WriteOutput(outputStream, MovePageRange(pdf, insertBeforePageNumber, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(byte[] pdf, Stream outputStream, int insertBeforePageNumber, PdfPageRange pageRange) {
        WriteOutput(outputStream, MovePageRange(pdf, insertBeforePageNumber, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(Stream inputStream, Stream outputStream, int insertBeforePageNumber, int firstPage, int lastPage) {
        WriteOutput(outputStream, MovePageRange(inputStream, insertBeforePageNumber, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(Stream inputStream, Stream outputStream, int insertBeforePageNumber, PdfPageRange pageRange) {
        WriteOutput(outputStream, MovePageRange(inputStream, insertBeforePageNumber, pageRange));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved.
    /// </summary>
    public static void MovePageRange(string inputPath, string outputPath, int insertBeforePageNumber, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(string inputPath, Stream outputStream, int insertBeforePageNumber, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved.
    /// </summary>
    public static void MovePageRange(string inputPath, string outputPath, int insertBeforePageNumber, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the inclusive one-based page range moved from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void MovePageRange(string inputPath, Stream outputStream, int insertBeforePageNumber, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved from a file path.
    /// </summary>
    public static byte[] MovePageRange(string inputPath, int insertBeforePageNumber, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF with the inclusive one-based page range moved from a file path.
    /// </summary>
    public static byte[] MovePageRange(string inputPath, int insertBeforePageNumber, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return MovePageRange(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRange);
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges moved before <paramref name="insertBeforePageNumber"/>.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static byte[] MovePageRanges(byte[] pdf, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        return MovePages(pdf, insertBeforePageNumber, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges moved from the current position of a readable stream.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static byte[] MovePageRanges(Stream stream, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        return MovePages(ReadStream(stream, nameof(stream)), insertBeforePageNumber, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges moved to <paramref name="outputStream"/>.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static void MovePageRanges(byte[] pdf, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, MovePageRanges(pdf, insertBeforePageNumber, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges moved from the current position of a readable stream to <paramref name="outputStream"/>.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static void MovePageRanges(Stream inputStream, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, MovePageRanges(inputStream, insertBeforePageNumber, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges moved.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static void MovePageRanges(string inputPath, string outputPath, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = MovePageRanges(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with the supplied inclusive one-based page ranges moved from a file path to <paramref name="outputStream"/>.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static void MovePageRanges(string inputPath, Stream outputStream, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, MovePageRanges(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with the supplied inclusive one-based page ranges moved from a file path.
    /// The moved pages keep their original relative order. Overlapping ranges are treated as one moved page set.
    /// </summary>
    public static byte[] MovePageRanges(string inputPath, int insertBeforePageNumber, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return MovePageRanges(File.ReadAllBytes(inputPath), insertBeforePageNumber, pageRanges);
    }
}
