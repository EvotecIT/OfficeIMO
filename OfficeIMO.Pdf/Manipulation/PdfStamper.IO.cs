using System.Threading;
using OfficeIMO.Core.Internal;
namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static byte[] ReadPath(string path, CancellationToken cancellationToken = default) =>
        PdfDocumentSource.FromPath(path, null, cancellationToken).Bytes;

    private static byte[] ReadStream(Stream stream, string paramName, PdfLoadOptions? readOptions = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        return PdfDocumentSource.FromRemainingStream(stream, readOptions, cancellationToken).Bytes;
    }

    private static byte[] ReadImageStream(Stream stream, string paramName, PdfImageStampOptions? options) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        return PdfImageInput.ReadRemainingStream(
            stream,
            (options ?? new PdfImageStampOptions()).MaximumEncodedImageBytes, options?.CancellationToken ?? default);
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
        cancellationToken.ThrowIfCancellationRequested();
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
