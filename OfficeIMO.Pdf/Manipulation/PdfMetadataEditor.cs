using OfficeIMO.Core.Internal;
using System.Threading;
namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party PDF metadata editing helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static partial class PdfMetadataEditor {
    /// <summary>
    /// Creates a new PDF with updated document metadata. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static byte[] UpdateMetadata(
        byte[] pdf,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        return UpdateMetadata(pdf, title, author, subject, keywords, readOptions: null);
    }

    internal static byte[] UpdateMetadata(
        byte[] pdf,
        string? title,
        string? author,
        string? subject,
        string? keywords,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.UpdateMetadata, readOptions, cancellationToken: cancellationToken);

        var document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        PdfMetadata existing = document.GetUncheckedMetadata(cancellationToken);
        var metadata = new PdfMetadata {
            Title = title ?? existing.Title,
            Author = author ?? existing.Author,
            Subject = subject ?? existing.Subject,
            Keywords = keywords ?? existing.Keywords,
            TrappingStatus = existing.TrappingStatus,
            PdfXVersion = existing.PdfXVersion,
            PdfXConformance = existing.PdfXConformance
        };
        metadata.CopySourceDatesFrom(existing);

        return RewriteWithMetadata(pdf, metadata, readOptions, cancellationToken);
    }

    /// <summary>
    /// Creates a new PDF with updated document metadata from the current position of a readable stream. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static byte[] UpdateMetadata(
        Stream stream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        return UpdateMetadata(stream, readOptions: null, CancellationToken.None, title, author, subject, keywords);
    }

    /// <summary>
    /// Writes a new PDF with updated document metadata to <paramref name="outputStream"/>. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static void UpdateMetadata(
        byte[] pdf,
        Stream outputStream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        WriteOutput(outputStream, UpdateMetadata(pdf, title, author, subject, keywords));
    }

    /// <summary>
    /// Writes a new PDF with updated document metadata from the current position of a readable stream to <paramref name="outputStream"/>. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static void UpdateMetadata(
        Stream inputStream,
        Stream outputStream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        WriteOutput(outputStream, UpdateMetadata(inputStream, title, author, subject, keywords));
    }

    /// <summary>
    /// Writes a new PDF with updated document metadata. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static void UpdateMetadata(
        string inputPath,
        string outputPath,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = UpdateMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, title, author, subject, keywords);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with updated document metadata from a file path to <paramref name="outputStream"/>. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static void UpdateMetadata(
        string inputPath,
        Stream outputStream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = UpdateMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, title, author, subject, keywords);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF with updated document metadata from a file path. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public static byte[] UpdateMetadataToBytes(
        string inputPath,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return UpdateMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, title, author, subject, keywords);
    }

    /// <summary>
    /// Creates a new PDF with exactly the supplied document metadata.
    /// </summary>
    public static byte[] ReplaceMetadata(byte[] pdf, PdfMetadata metadata) {
        return ReplaceMetadata(pdf, metadata, readOptions: null);
    }

    internal static byte[] ReplaceMetadata(byte[] pdf, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(metadata, nameof(metadata));
        cancellationToken.ThrowIfCancellationRequested();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.UpdateMetadata, readOptions, cancellationToken: cancellationToken);

        return RewriteWithMetadata(pdf, metadata, readOptions, cancellationToken);
    }

    /// <summary>
    /// Creates a new PDF with exactly the supplied document metadata from the current position of a readable stream.
    /// </summary>
    public static byte[] ReplaceMetadata(Stream stream, PdfMetadata metadata) {
        return ReplaceMetadata(stream, metadata, readOptions: null, CancellationToken.None);
    }

    /// <summary>
    /// Writes a new PDF with exactly the supplied document metadata to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReplaceMetadata(byte[] pdf, Stream outputStream, PdfMetadata metadata) {
        WriteOutput(outputStream, ReplaceMetadata(pdf, metadata));
    }

    /// <summary>
    /// Writes a new PDF with exactly the supplied document metadata from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReplaceMetadata(Stream inputStream, Stream outputStream, PdfMetadata metadata) {
        WriteOutput(outputStream, ReplaceMetadata(inputStream, metadata));
    }

    /// <summary>
    /// Writes a new PDF with exactly the supplied document metadata.
    /// </summary>
    public static void ReplaceMetadata(string inputPath, string outputPath, PdfMetadata metadata) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));
        Guard.NotNull(metadata, nameof(metadata));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ReplaceMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, metadata);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF with exactly the supplied document metadata from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ReplaceMetadata(string inputPath, Stream outputStream, PdfMetadata metadata) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);
        Guard.NotNull(metadata, nameof(metadata));

        var bytes = ReplaceMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, metadata);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF with exactly the supplied document metadata from a file path.
    /// </summary>
    public static byte[] ReplaceMetadataToBytes(string inputPath, PdfMetadata metadata) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(metadata, nameof(metadata));
        return ReplaceMetadata(PdfDocumentSource.FromPath(inputPath, null).Bytes, metadata);
    }

    private static byte[] RewriteWithMetadata(byte[] pdf, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        var document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var pageObjectNumbers = new int[document.Pages.Count];
        for (int index = 0; index < pageObjectNumbers.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            pageObjectNumbers[index] = document.Pages[index].ObjectNumber;
        }
        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, metadata, pageObjectNumbers,
            catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw, document, cancellationToken),
            fileVersion: fileVersion, cancellationToken: cancellationToken);
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes, CancellationToken cancellationToken = default) {
        ValidateWritableOutputStream(outputStream);
        OfficeStreamWriter.CopyBytes(outputStream, bytes, cancellationToken);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static void WriteOutput(string outputPath, byte[] bytes, CancellationToken cancellationToken = default) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullPath, stream => OfficeStreamWriter.CopyBytes(stream, bytes, cancellationToken), cancellationToken);
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
