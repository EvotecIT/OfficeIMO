using OfficeIMO.Drawing.Internal;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Extracts image XObjects from PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static class PdfImageExtractor {
    /// <summary>
    /// Extracts image XObjects from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImages(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractImages(PdfReadDocument.Open(pdf));
    }

    /// <summary>
    /// Extracts image XObject placement invocations from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacements(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractImagePlacements(PdfReadDocument.Open(pdf));
    }

    /// <summary>
    /// Extracts image XObjects from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImages(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractImages(File.ReadAllBytes(path));
    }

    /// <summary>
    /// Extracts image XObject placement invocations from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacements(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractImagePlacements(PdfReadDocument.Open(path));
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImagesByPageRanges(string path, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractImagesByPageRanges(PdfReadDocument.Open(path), pageRanges);
    }

    /// <summary>
    /// Extracts image XObject placement invocations from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacementsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractImagePlacementsByPageRanges(PdfReadDocument.Open(path), pageRanges);
    }

    /// <summary>
    /// Extracts image XObjects from all pages in page order from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImages(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return ExtractImages(buffer.ToArray());
    }

    /// <summary>
    /// Extracts image XObject placement invocations from all pages in page order from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacements(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return ExtractImagePlacements(buffer.ToArray());
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImagesByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return ExtractImagesByPageRanges(buffer.ToArray(), pageRanges);
    }

    /// <summary>
    /// Extracts image XObject placement invocations from the supplied inclusive one-based page ranges from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacementsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return ExtractImagePlacementsByPageRanges(buffer.ToArray(), pageRanges);
    }

    /// <summary>
    /// Extracts image XObjects from a PDF path and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImages(string inputPath, string outputDirectory) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImages(inputPath);
        return WriteImageFiles(images, fullOutputDirectory, inputPath);
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges in a PDF path and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImagesByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImagesByPageRanges(inputPath, pageRanges);
        return WriteImageFiles(images, fullOutputDirectory, inputPath);
    }

    /// <summary>
    /// Extracts image XObjects from the current position of a readable stream and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImages(Stream stream, string outputDirectory, string baseName = "image") {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImages(stream);
        return WriteImageFiles(images, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges from the current position of a readable stream and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImagesByPageRanges(Stream stream, string outputDirectory, string baseName = "image", params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImagesByPageRanges(stream, pageRanges);
        return WriteImageFiles(images, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Extracts image XObjects from bytes and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImages(byte[] pdf, string outputDirectory, string baseName = "image") {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImages(pdf);
        return WriteImageFiles(images, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges from bytes and writes them to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> ExtractImagesByPageRanges(byte[] pdf, string outputDirectory, string baseName = "image", params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var images = ExtractImagesByPageRanges(pdf, pageRanges);
        return WriteImageFiles(images, fullOutputDirectory, baseName);
    }

    private static List<string> WriteImageFiles(IReadOnlyList<PdfExtractedImage> images, string fullOutputDirectory, string baseName) {
        string safeBaseName = GetSafeBaseName(baseName, "image");

        var paths = new List<string>(images.Count);
        for (int i = 0; i < images.Count; i++) {
            var image = images[i];
            string extension = string.IsNullOrWhiteSpace(image.FileExtension) ? "bin" : image.FileExtension!.TrimStart('.');
            string outputPath = Path.Combine(
                fullOutputDirectory,
                safeBaseName +
                "-page-" + image.PageNumber.ToString("0000", CultureInfo.InvariantCulture) +
                "-image-" + (i + 1).ToString("0000", CultureInfo.InvariantCulture) +
                "." + extension);

            OfficeFileCommit.WriteAllBytes(outputPath, image.Bytes);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static string GetSafeBaseName(string? baseName, string fallback) {
        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        return string.IsNullOrWhiteSpace(safeBaseName) ? fallback : safeBaseName;
    }

    private static string ValidateOutputDirectory(string outputDirectory) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("Output directory cannot be empty or whitespace.", nameof(outputDirectory));
        }

        string fullOutputDirectory;
        try {
            fullOutputDirectory = Path.GetFullPath(outputDirectory);
        } catch (Exception ex) {
            throw new ArgumentException("Output directory is invalid.", nameof(outputDirectory), ex);
        }

        if (File.Exists(fullOutputDirectory)) {
            throw new ArgumentException("Output directory refers to a file; a directory path is required.", nameof(outputDirectory));
        }

        Directory.CreateDirectory(fullOutputDirectory);
        return fullOutputDirectory;
    }

    /// <summary>
    /// Extracts image XObjects from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImages(PdfReadDocument document) {
        Guard.NotNull(document, nameof(document));

        var images = new List<PdfExtractedImage>();
        for (int i = 0; i < document.Pages.Count; i++) {
            images.AddRange(document.Pages[i].GetImages(i + 1));
        }

        return images;
    }

    /// <summary>
    /// Extracts image XObject placement invocations from all pages in page order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacements(PdfReadDocument document) {
        Guard.NotNull(document, nameof(document));

        var placements = new List<PdfImagePlacement>();
        for (int i = 0; i < document.Pages.Count; i++) {
            placements.AddRange(document.Pages[i].GetImagePlacements(i + 1));
        }

        return placements;
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImagesByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractImagesByPageRanges(PdfReadDocument.Open(pdf), pageRanges);
    }

    /// <summary>
    /// Extracts image XObject placement invocations from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacementsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractImagePlacementsByPageRanges(PdfReadDocument.Open(pdf), pageRanges);
    }

    /// <summary>
    /// Extracts image XObjects from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfExtractedImage> ExtractImagesByPageRanges(PdfReadDocument document, params PdfPageRange[] pageRanges) {
        Guard.NotNull(document, nameof(document));
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
        var images = new List<PdfExtractedImage>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            images.AddRange(document.Pages[pageNumber - 1].GetImages(pageNumber));
        }

        return images;
    }

    /// <summary>
    /// Extracts image XObject placement invocations from the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static IReadOnlyList<PdfImagePlacement> ExtractImagePlacementsByPageRanges(PdfReadDocument document, params PdfPageRange[] pageRanges) {
        Guard.NotNull(document, nameof(document));
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
        var placements = new List<PdfImagePlacement>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            placements.AddRange(document.Pages[pageNumber - 1].GetImagePlacements(pageNumber));
        }

        return placements;
    }

}
