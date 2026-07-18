using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party page extraction and simple splitting for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static partial class PdfPageExtractor {
    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, params int[] pageNumbers) {
        return ExtractPages(pdf, (IEnumerable<int>)pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPages(Stream stream, params int[] pageNumbers) {
        return ExtractPages(stream, (IEnumerable<int>)pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, IEnumerable<int> pageNumbers) {
        return ExtractPages(pdf, pageNumbers, options: null);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order, using read options for password-protected sources.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, PdfReadOptions? options, params int[] pageNumbers) {
        return ExtractPages(pdf, pageNumbers, options);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order, using read options for password-protected sources.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, IEnumerable<int> pageNumbers, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ExtractPages, options);

        var selected = pageNumbers.ToArray();
        if (selected.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        ValidatePageNumbers(selected, document.Pages.Count, nameof(pageNumbers));

        var pageObjectNumbers = selected.Select(pageNumber => document.Pages[pageNumber - 1].ObjectNumber).ToArray();
        PdfFileVersion fileVersion = GetSourceFileVersion(pdf);
        return ExtractPages(objects, document.Metadata, pageObjectNumbers, catalogState: ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPages(Stream stream, IEnumerable<int> pageNumbers) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        return ExtractPages(ReadStream(stream, nameof(stream)), pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(byte[] pdf, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ExtractPages(pdf, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(Stream inputStream, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ExtractPages(inputStream, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static byte[] ExtractPageRange(byte[] pdf, int firstPage, int lastPage) {
        if (firstPage > lastPage) {
            throw new ArgumentOutOfRangeException(nameof(lastPage), "Last page must be greater than or equal to first page.");
        }

        return ExtractPages(pdf, Enumerable.Range(firstPage, lastPage - firstPage + 1));
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static byte[] ExtractPageRange(byte[] pdf, PdfPageRange pageRange) {
        return ExtractPages(pdf, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRange(Stream stream, int firstPage, int lastPage) {
        return ExtractPageRange(ReadStream(stream, nameof(stream)), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRange(Stream stream, PdfPageRange pageRange) {
        return ExtractPageRange(ReadStream(stream, nameof(stream)), pageRange);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(byte[] pdf, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, ExtractPageRange(pdf, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(byte[] pdf, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, ExtractPageRange(pdf, pageRange));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(Stream inputStream, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, ExtractPageRange(inputStream, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(Stream inputStream, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, ExtractPageRange(inputStream, pageRange));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static void ExtractPageRange(string inputPath, string outputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(string inputPath, Stream outputStream, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static void ExtractPageRange(string inputPath, string outputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(string inputPath, Stream outputStream, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from a file path.
    /// </summary>
    public static byte[] ExtractPageRange(string inputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from a file path.
    /// </summary>
    public static byte[] ExtractPageRange(string inputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges) {
        return ExtractPageRanges(pdf, pageRanges, options: null);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order, using read options for password-protected sources.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, PdfReadOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges, options);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order, using read options for password-protected sources.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ExtractPages, options);

        var ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        ValidatePageRanges(ranges, document.Pages.Count, nameof(pageRanges));

        var pageObjectNumbers = new List<int>(ranges.Sum(range => range.PageCount));
        foreach (var range in ranges) {
            for (int pageNumber = range.FirstPage; pageNumber <= range.LastPage; pageNumber++) {
                pageObjectNumbers.Add(document.Pages[pageNumber - 1].ObjectNumber);
            }
        }

        PdfFileVersion fileVersion = GetSourceFileVersion(pdf);
        return ExtractPages(objects, document.Metadata, pageObjectNumbers.ToArray(), catalogState: ExtractCatalogRewriteState(objects, trailerRaw), fileVersion: fileVersion);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRanges(Stream stream, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return ExtractPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ExtractPageRanges(pdf, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ExtractPageRanges(inputStream, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static void ExtractPageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from a file path.
    /// </summary>
    public static byte[] ExtractPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(byte[] pdf) {
        return SplitPages(pdf, options: null);
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page, using read options for password-protected sources.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(byte[] pdf, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ExtractPages, options);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        var catalogState = ExtractCatalogRewriteState(objects, trailerRaw);
        PdfFileVersion fileVersion = GetSourceFileVersion(pdf);
        var result = new List<byte[]>(document.Pages.Count);

        foreach (var page in document.Pages) {
            result.Add(ExtractPages(objects, document.Metadata, new[] { page.ObjectNumber }, catalogState: catalogState, fileVersion: fileVersion));
        }

        return result;
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(Stream stream) {
        return SplitPages(ReadStream(stream, nameof(stream)));
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return SplitPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges) {
        return SplitPageRanges(pdf, pageRanges, options: null);
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range, using read options for password-protected sources.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, PdfReadOptions? options, params PdfPageRange[] pageRanges) {
        return SplitPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges, options);
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range, using read options for password-protected sources.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ExtractPages, options);

        var ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        ValidatePageRanges(ranges, document.Pages.Count, nameof(pageRanges));
        var catalogState = ExtractCatalogRewriteState(objects, trailerRaw);
        PdfFileVersion fileVersion = GetSourceFileVersion(pdf);
        var result = new List<byte[]>(ranges.Length);

        foreach (var range in ranges) {
            int[] pageObjectNumbers = Enumerable
                .Range(range.FirstPage, range.PageCount)
                .Select(pageNumber => document.Pages[pageNumber - 1].ObjectNumber)
                .ToArray();
            result.Add(ExtractPages(objects, document.Metadata, pageObjectNumbers, catalogState: catalogState, fileVersion: fileVersion));
        }

        return result;
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return SplitPageRanges(stream, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(Stream stream, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return SplitPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPages(string inputPath, string outputDirectory) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPages(File.ReadAllBytes(inputPath));
        string baseName = Path.GetFileNameWithoutExtension(inputPath);
        return WriteSplitPages(pages, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Splits a PDF from a file path into one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return SplitPages(File.ReadAllBytes(inputPath));
    }

    /// <summary>
    /// Splits a PDF from a file path into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        return SplitPageRanges(File.ReadAllBytes(inputPath), ranges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream and writes one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<string> SplitPages(Stream stream, string outputDirectory, string baseName = "page") {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPages(stream);
        return WriteSplitPages(pages, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Splits a PDF from a file path into one output PDF per inclusive one-based page range and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPageRanges(File.ReadAllBytes(inputPath), ranges);
        string baseName = Path.GetFileNameWithoutExtension(inputPath);
        return WriteSplitPageRanges(pages, fullOutputDirectory, baseName, ranges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPageRanges(Stream stream, string outputDirectory, string baseName, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPageRanges(stream, ranges);
        return WriteSplitPageRanges(pages, fullOutputDirectory, baseName, ranges);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers.
    /// </summary>
    public static void ExtractPages(string inputPath, string outputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers from a file path.
    /// </summary>
    public static byte[] ExtractPages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
    }

}
