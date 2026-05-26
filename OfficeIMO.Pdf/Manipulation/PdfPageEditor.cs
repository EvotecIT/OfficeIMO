using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party page editing helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
public static class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the specified one-based pages removed.
    /// </summary>
    public static byte[] DeletePages(byte[] pdf, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, remaining.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects));
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

    /// <summary>
    /// Creates a new PDF with the specified one-based pages duplicated immediately after each selected source page.
    /// Repeated selections create repeated page copies.
    /// </summary>
    public static byte[] DuplicatePages(byte[] pdf, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects));
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

    /// <summary>
    /// Creates a new PDF with the specified one-based pages moved before <paramref name="insertBeforePageNumber"/>.
    /// The moved pages keep their original relative order. Use page count + 1 to move pages to the end.
    /// </summary>
    public static byte[] MovePages(byte[] pdf, int insertBeforePageNumber, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered.ToArray(), catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects));
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

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public static byte[] ReorderPages(byte[] pdf, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        ValidateReorderPageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers));

        var ordered = new int[pageNumbers.Length];
        for (int i = 0; i < pageNumbers.Length; i++) {
            ordered[i] = document.Pages[pageNumbers[i] - 1].ObjectNumber;
        }

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, ordered, catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects));
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

    /// <summary>
    /// Creates a new PDF with the selected pages rotated to the specified degrees. If no page numbers are supplied, all pages are rotated.
    /// </summary>
    public static byte[] RotatePages(byte[] pdf, int rotationDegrees, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        int normalizedRotation = NormalizeRotation(rotationDegrees);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects));
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

    private static void ValidatePageNumbers(int[] pageNumbers, int pageCount, string paramName, bool allowDuplicates = false) {
        var seen = new HashSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1 || pageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(paramName, "Page number " + pageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }

            if (!allowDuplicates && !seen.Add(pageNumber)) {
                throw new ArgumentException("Duplicate page selections are not supported.", paramName);
            }
        }
    }

    private static void ValidateReorderPageNumbers(int[] pageNumbers, int pageCount, string paramName) {
        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", paramName);
        }

        if (pageNumbers.Length != pageCount) {
            throw new ArgumentException("Reorder must include every page exactly once.", paramName);
        }

        ValidatePageNumbers(pageNumbers, pageCount, paramName);
    }

    private static int[] ExpandPageRanges(PdfPageRange[] pageRanges, string paramName) {
        Guard.NotNull(pageRanges, paramName);
        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", paramName);
        }

        var pages = new List<int>();
        for (int i = 0; i < pageRanges.Length; i++) {
            pages.AddRange(pageRanges[i].ToPageNumbers());
        }

        return pages.ToArray();
    }

    private static int[] ExpandPageRangesDistinct(PdfPageRange[] pageRanges, string paramName) {
        int[] pages = ExpandPageRanges(pageRanges, paramName);
        var seen = new HashSet<int>();
        var distinct = new List<int>(pages.Length);
        for (int i = 0; i < pages.Length; i++) {
            if (seen.Add(pages[i])) {
                distinct.Add(pages[i]);
            }
        }

        return distinct.ToArray();
    }

    private static void ValidateMoveInsertBeforePageNumber(int insertBeforePageNumber, int pageCount) {
        if (insertBeforePageNumber < 1 || insertBeforePageNumber > pageCount + 1) {
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber), "Insert-before page must be in the document page range 1-" + (pageCount + 1).ToString(CultureInfo.InvariantCulture) + ".");
        }
    }

    private static int[] BuildInclusivePageRange(int firstPage, int lastPage, string lastPageParamName) {
        if (firstPage > lastPage) {
            throw new ArgumentOutOfRangeException(lastPageParamName, "Last page must be greater than or equal to first page.");
        }

        return Enumerable.Range(firstPage, lastPage - firstPage + 1).ToArray();
    }

    private static int NormalizeRotation(int rotationDegrees) {
        if (rotationDegrees % 90 != 0) {
            throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be a multiple of 90 degrees.");
        }

        int normalized = rotationDegrees % 360;
        if (normalized < 0) {
            normalized += 360;
        }

        return normalized;
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
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, bytes);
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
