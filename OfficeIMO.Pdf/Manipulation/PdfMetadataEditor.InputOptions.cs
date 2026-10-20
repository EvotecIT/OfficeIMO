using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfMetadataEditor {
    /// <summary>Finishes bounded metadata updating before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void UpdateMetadata(byte[] pdf, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = UpdateMetadata(pdf, title, author, subject, keywords, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Updates metadata from a bounded remaining stream snapshot without taking ownership of the stream.</summary>
    internal static byte[] UpdateMetadata(Stream input, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(input, readOptions, cancellationToken);
        return UpdateMetadata(source.Bytes, title, author, subject, keywords, source.Options, cancellationToken);
    }

    /// <summary>Updates metadata from a bounded file snapshot with cooperative cancellation.</summary>
    internal static byte[] UpdateMetadataToBytes(string inputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        PdfDocumentSource source = PdfDocumentSource.FromPath(inputPath, readOptions, cancellationToken);
        return UpdateMetadata(source.Bytes, title, author, subject, keywords, source.Options, cancellationToken);
    }

    /// <summary>Finishes metadata rewriting before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void UpdateMetadata(Stream input, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = UpdateMetadata(input, readOptions, cancellationToken, title, author, subject, keywords);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Finishes metadata rewriting before checking cancellation and atomically publishing the destination file.</summary>
    internal static void UpdateMetadata(string inputPath, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        string destination = ValidateOutputPath(outputPath);
        byte[] rewritten = UpdateMetadataToBytes(inputPath, readOptions, cancellationToken, title, author, subject, keywords);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, rewritten, cancellationToken);
    }

    /// <summary>Reads bounded file input and finishes metadata updating before changing a caller-owned output stream.</summary>
    internal static void UpdateMetadata(string inputPath, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = UpdateMetadataToBytes(inputPath, readOptions, cancellationToken, title, author, subject, keywords);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Finishes bounded metadata replacement before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void ReplaceMetadata(byte[] pdf, Stream output, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = ReplaceMetadata(pdf, metadata, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Replaces metadata from a bounded remaining stream snapshot while retaining caller stream ownership.</summary>
    internal static byte[] ReplaceMetadata(Stream input, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        Guard.NotNull(metadata, nameof(metadata));
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(input, readOptions, cancellationToken);
        return ReplaceMetadata(source.Bytes, metadata, source.Options, cancellationToken);
    }

    /// <summary>Replaces metadata from a bounded file snapshot with cooperative cancellation.</summary>
    internal static byte[] ReplaceMetadataToBytes(string inputPath, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        Guard.NotNull(metadata, nameof(metadata));
        PdfDocumentSource source = PdfDocumentSource.FromPath(inputPath, readOptions, cancellationToken);
        return ReplaceMetadata(source.Bytes, metadata, source.Options, cancellationToken);
    }

    /// <summary>Finishes metadata replacement before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void ReplaceMetadata(Stream input, Stream output, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = ReplaceMetadata(input, metadata, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Finishes metadata replacement before checking cancellation and atomically publishing the destination file.</summary>
    internal static void ReplaceMetadata(string inputPath, string outputPath, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        string destination = ValidateOutputPath(outputPath);
        byte[] rewritten = ReplaceMetadataToBytes(inputPath, metadata, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, rewritten, cancellationToken);
    }

    /// <summary>Reads bounded file input and finishes metadata replacement before changing a caller-owned output stream.</summary>
    internal static void ReplaceMetadata(string inputPath, Stream output, PdfMetadata metadata, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = ReplaceMetadataToBytes(inputPath, metadata, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Finishes bounded Info/XMP synchronization before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void SynchronizeMetadata(byte[] pdf, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = SynchronizeMetadata(pdf, title, author, subject, keywords, createXmpMetadata, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Synchronizes Info and XMP using bounded remaining input, retaining caller stream ownership.</summary>
    internal static byte[] SynchronizeMetadata(Stream input, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(input, readOptions, cancellationToken);
        return SynchronizeMetadata(source.Bytes, title, author, subject, keywords, createXmpMetadata, source.Options, cancellationToken);
    }

    /// <summary>Synchronizes Info and XMP using bounded file input and cooperative cancellation.</summary>
    internal static byte[] SynchronizeMetadataToBytes(string inputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        PdfDocumentSource source = PdfDocumentSource.FromPath(inputPath, readOptions, cancellationToken);
        return SynchronizeMetadata(source.Bytes, title, author, subject, keywords, createXmpMetadata, source.Options, cancellationToken);
    }

    /// <summary>Finishes Info/XMP synchronization before checking cancellation and changing a caller-owned output stream.</summary>
    internal static void SynchronizeMetadata(Stream input, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = SynchronizeMetadata(input, readOptions, cancellationToken, title, author, subject, keywords, createXmpMetadata);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }

    /// <summary>Finishes Info/XMP synchronization before checking cancellation and atomically publishing the destination file.</summary>
    internal static void SynchronizeMetadata(string inputPath, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        string destination = ValidateOutputPath(outputPath);
        byte[] rewritten = SynchronizeMetadataToBytes(inputPath, readOptions, cancellationToken, title, author, subject, keywords, createXmpMetadata);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, rewritten, cancellationToken);
    }

    /// <summary>Reads bounded file input and finishes Info/XMP synchronization before changing a caller-owned output stream.</summary>
    internal static void SynchronizeMetadata(string inputPath, Stream output, PdfLoadOptions? readOptions, CancellationToken cancellationToken,
        string? title = null, string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true) {
        ValidateWritableOutputStream(output);
        byte[] rewritten = SynchronizeMetadataToBytes(inputPath, readOptions, cancellationToken, title, author, subject, keywords, createXmpMetadata);
        cancellationToken.ThrowIfCancellationRequested();
        WriteOutput(output, rewritten, cancellationToken);
    }
}
