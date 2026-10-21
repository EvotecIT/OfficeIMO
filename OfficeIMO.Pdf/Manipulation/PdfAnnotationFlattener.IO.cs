using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    /// <summary>Reads remaining stream content under caller-selected limits without closing the input.</summary>
    internal static byte[] FlattenVisualAnnotations(Stream stream, PdfAnnotationFlattenOptions? options, PdfLoadOptions? readOptions) {
        var cancellationToken = options?.CancellationToken ?? default;
        byte[] source = PdfDocumentSource.FromRemainingStream(stream, readOptions, cancellationToken).Bytes;
        return FlattenVisualAnnotations(source, options, readOptions);
    }

    /// <summary>Prepares flattened bytes before checking cancellation and changing the caller's output stream.</summary>
    internal static void FlattenVisualAnnotations(Stream inputStream, Stream outputStream, PdfAnnotationFlattenOptions? options,
        PdfLoadOptions? readOptions) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = FlattenVisualAnnotations(inputStream, options, readOptions);
        (options?.CancellationToken ?? default).ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options?.CancellationToken ?? default);
    }

    /// <summary>Reads one bounded file snapshot with cancellation before parsing and rewriting.</summary>
    internal static byte[] FlattenVisualAnnotationsToBytes(string inputPath, PdfAnnotationFlattenOptions? options, PdfLoadOptions? readOptions) {
        var cancellationToken = options?.CancellationToken ?? default;
        byte[] source = PdfDocumentSource.FromPath(inputPath, readOptions, cancellationToken).Bytes;
        return FlattenVisualAnnotations(source, options, readOptions);
    }

    private static byte[] ReadStream(Stream stream) {
        return PdfDocumentSource.FromRemainingStream(stream, null).Bytes;
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes, CancellationToken cancellationToken = default) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }

        OfficeStreamWriter.CopyBytes(outputStream, bytes, cancellationToken);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
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
