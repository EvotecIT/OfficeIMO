using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    /// <summary>Extracts selected pages from a bounded file snapshot with cooperative cancellation.</summary>
    internal static byte[] ExtractPages(string inputPath, IEnumerable<int> pageNumbers, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPages(ReadPath(inputPath, options, cancellationToken), pageNumbers, options, documentFactory: null, cancellationToken);

    /// <summary>Extracts selected pages and checks cancellation before changing the caller's output stream.</summary>
    internal static void ExtractPages(byte[] pdf, Stream outputStream, IEnumerable<int> pageNumbers, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = ExtractPages(pdf, pageNumbers, options, documentFactory: null, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, cancellationToken);
    }

    /// <summary>Extracts selected pages and checks cancellation before publishing the destination file.</summary>
    internal static void ExtractPages(string inputPath, string outputPath, IEnumerable<int> pageNumbers, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] output = ExtractPages(inputPath, pageNumbers, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(fullOutputPath, output, cancellationToken);
    }

    /// <summary>Extracts selected pages from a file without closing or changing output until extraction succeeds.</summary>
    internal static void ExtractPages(string inputPath, Stream outputStream, IEnumerable<int> pageNumbers, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = ExtractPages(inputPath, pageNumbers, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, cancellationToken);
    }

    /// <summary>Extracts one inclusive page range with caller-selected admission limits and cancellation.</summary>
    internal static byte[] ExtractPageRange(byte[] pdf, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(pdf, new[] { PdfPageRange.From(firstPage, lastPage) }, options, cancellationToken: cancellationToken);

    /// <summary>Extracts one inclusive page range with caller-selected admission limits and cancellation.</summary>
    internal static byte[] ExtractPageRange(byte[] pdf, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(pdf, new[] { pageRange }, options, cancellationToken: cancellationToken);

    /// <summary>Reads the remaining stream without closing it and extracts one bounded page range.</summary>
    internal static byte[] ExtractPageRange(Stream stream, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(stream, new[] { PdfPageRange.From(firstPage, lastPage) }, options, cancellationToken);

    /// <summary>Reads the remaining stream without closing it and extracts one bounded page range.</summary>
    internal static byte[] ExtractPageRange(Stream stream, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(stream, new[] { pageRange }, options, cancellationToken);

    /// <summary>Reads one bounded file snapshot and extracts one inclusive page range.</summary>
    internal static byte[] ExtractPageRange(string inputPath, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(inputPath, new[] { PdfPageRange.From(firstPage, lastPage) }, options, cancellationToken);

    /// <summary>Reads one bounded file snapshot and extracts one inclusive page range.</summary>
    internal static byte[] ExtractPageRange(string inputPath, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(inputPath, new[] { pageRange }, options, cancellationToken);

    /// <summary>Extracts one inclusive range before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void ExtractPageRange(byte[] pdf, Stream outputStream, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(pdf, outputStream, new[] { pageRange }, options, cancellationToken);

    /// <summary>Extracts one inclusive range with explicit limits and cancellation before output publication.</summary>
    internal static void ExtractPageRange(byte[] pdf, Stream outputStream, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRange(pdf, outputStream, PdfPageRange.From(firstPage, lastPage), options, cancellationToken);

    /// <summary>Consumes bounded remaining input and finishes range extraction before changing the caller-owned output.</summary>
    internal static void ExtractPageRange(Stream inputStream, Stream outputStream, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(inputStream, outputStream, new[] { pageRange }, options, cancellationToken);

    /// <summary>Consumes bounded remaining input and finishes range extraction before changing the caller-owned output.</summary>
    internal static void ExtractPageRange(Stream inputStream, Stream outputStream, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRange(inputStream, outputStream, PdfPageRange.From(firstPage, lastPage), options, cancellationToken);

    /// <summary>Finishes bounded file range extraction before checking cancellation and atomically publishing the destination.</summary>
    internal static void ExtractPageRange(string inputPath, string outputPath, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(inputPath, outputPath, new[] { pageRange }, options, cancellationToken);

    /// <summary>Finishes bounded file range extraction before checking cancellation and atomically publishing the destination.</summary>
    internal static void ExtractPageRange(string inputPath, string outputPath, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRange(inputPath, outputPath, PdfPageRange.From(firstPage, lastPage), options, cancellationToken);

    /// <summary>Finishes bounded file range extraction before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void ExtractPageRange(string inputPath, Stream outputStream, PdfPageRange pageRange, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(inputPath, outputStream, new[] { pageRange }, options, cancellationToken);

    /// <summary>Finishes bounded file range extraction before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void ExtractPageRange(string inputPath, Stream outputStream, int firstPage, int lastPage, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRange(inputPath, outputStream, PdfPageRange.From(firstPage, lastPage), options, cancellationToken);

    /// <summary>Extracts ordered ranges from a bounded file snapshot with cooperative cancellation.</summary>
    internal static byte[] ExtractPageRanges(string inputPath, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        ExtractPageRanges(ReadPath(inputPath, options, cancellationToken), pageRanges, options, cancellationToken: cancellationToken);

    /// <summary>Extracts ordered ranges and checks cancellation before changing the caller's output stream.</summary>
    internal static void ExtractPageRanges(byte[] pdf, Stream outputStream, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = ExtractPageRanges(pdf, pageRanges, options, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, cancellationToken);
    }

    /// <summary>Consumes remaining input without closing either stream; admission and cancellation complete before output starts.</summary>
    internal static void ExtractPageRanges(Stream inputStream, Stream outputStream, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = ExtractPageRanges(inputStream, pageRanges, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, cancellationToken);
    }

    /// <summary>Extracts ordered ranges from a file and checks cancellation before destination publication.</summary>
    internal static void ExtractPageRanges(string inputPath, string outputPath, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] output = ExtractPageRanges(inputPath, pageRanges, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(fullOutputPath, output, cancellationToken);
    }

    /// <summary>Finishes bounded file range extraction before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void ExtractPageRanges(string inputPath, Stream outputStream, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = ExtractPageRanges(inputPath, pageRanges, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, cancellationToken);
    }

    /// <summary>Splits a bounded file snapshot into caller-ordered ranges with cancellation.</summary>
    internal static IReadOnlyList<byte[]> SplitPageRanges(string inputPath, IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        SplitPageRanges(ReadPath(inputPath, options, cancellationToken), pageRanges, options, cancellationToken: cancellationToken);

    /// <summary>Prepares every split part before checking cancellation and starting file publication.</summary>
    /// <remarks>Cancellation is observed before publishing the prepared set. Once file publication begins, the set is written without further cancellation.</remarks>
    internal static IReadOnlyList<string> SplitPages(string inputPath, string outputDirectory, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        IReadOnlyList<byte[]> pages = SplitPages(inputPath, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return WriteSplitPages(pages, fullOutputDirectory, Path.GetFileNameWithoutExtension(inputPath));
    }

    /// <summary>Consumes remaining input without closing it and prepares all split parts before file publication.</summary>
    /// <remarks>Cancellation is observed before publishing the prepared set. Once file publication begins, the set is written without further cancellation.</remarks>
    internal static IReadOnlyList<string> SplitPages(Stream stream, string outputDirectory, string baseName, PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        IReadOnlyList<byte[]> pages = SplitPages(stream, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return WriteSplitPages(pages, fullOutputDirectory, baseName);
    }

    /// <summary>Prepares every range part before checking cancellation and starting file publication.</summary>
    /// <remarks>Cancellation is observed before publishing the prepared set. Once file publication begins, the set is written without further cancellation.</remarks>
    internal static IReadOnlyList<string> SplitPageRanges(string inputPath, string outputDirectory, IEnumerable<PdfPageRange> pageRanges,
        PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        PdfPageRange[] ranges = CaptureRanges(pageRanges, cancellationToken);
        IReadOnlyList<byte[]> pages = SplitPageRanges(inputPath, ranges, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return WriteSplitPageRanges(pages, fullOutputDirectory, Path.GetFileNameWithoutExtension(inputPath), ranges);
    }

    /// <summary>Consumes remaining input and prepares every range part before checking cancellation and publishing files.</summary>
    /// <remarks>Cancellation is observed before publishing the prepared set. Once file publication begins, the set is written without further cancellation.</remarks>
    internal static IReadOnlyList<string> SplitPageRanges(Stream stream, string outputDirectory, string baseName,
        IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        PdfPageRange[] ranges = CaptureRanges(pageRanges, cancellationToken);
        IReadOnlyList<byte[]> pages = SplitPageRanges(stream, ranges, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return WriteSplitPageRanges(pages, fullOutputDirectory, baseName, ranges);
    }

    private static PdfPageRange[] CaptureRanges(IEnumerable<PdfPageRange> pageRanges, CancellationToken cancellationToken) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        cancellationToken.ThrowIfCancellationRequested();
        var ranges = new List<PdfPageRange>();
        foreach (PdfPageRange range in pageRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            ranges.Add(range);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return ranges.ToArray();
    }
}
