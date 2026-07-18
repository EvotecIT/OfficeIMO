using OfficeIMO.Drawing.Internal;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
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
    
    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }
    
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
    
    private static List<string> WriteSplitPages(IReadOnlyList<byte[]> pages, string outputDirectory, string? baseName) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
    
        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(safeBaseName)) {
            safeBaseName = "page";
        }
    
        var paths = new List<string>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string outputPath = Path.Combine(fullOutputDirectory, safeBaseName + "-page-" + (i + 1).ToString("0000", CultureInfo.InvariantCulture) + ".pdf");
            OfficeFileCommit.WriteAllBytes(outputPath, pages[i]);
            paths.Add(outputPath);
        }
    
        return paths;
    }
    
    private static List<string> WriteSplitPageRanges(IReadOnlyList<byte[]> pages, string outputDirectory, string? baseName, PdfPageRange[] ranges) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
    
        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(safeBaseName)) {
            safeBaseName = "page";
        }
    
        var paths = new List<string>(pages.Count);
        var rangeOccurrences = new Dictionary<PdfPageRange, int>();
        for (int i = 0; i < pages.Count; i++) {
            var range = ranges[i];
            rangeOccurrences.TryGetValue(range, out int occurrence);
            occurrence++;
            rangeOccurrences[range] = occurrence;
    
            string outputPath = Path.Combine(
                fullOutputDirectory,
                safeBaseName + "-pages-" +
                range.FirstPage.ToString("0000", CultureInfo.InvariantCulture) + "-" +
                range.LastPage.ToString("0000", CultureInfo.InvariantCulture) +
                (occurrence <= 1 ? string.Empty : "-occurrence-" + occurrence.ToString("0000", CultureInfo.InvariantCulture)) +
                ".pdf");
            OfficeFileCommit.WriteAllBytes(outputPath, pages[i]);
            paths.Add(outputPath);
        }
    
        return paths;
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
    
    private static PdfPageRange[] ValidatePageRangeArguments(PdfPageRange[]? pageRanges, string paramName) {
        if (pageRanges is null) {
            throw new ArgumentNullException(paramName);
        }
    
        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", paramName);
        }
    
        return (PdfPageRange[])pageRanges.Clone();
    }
    
    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
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
