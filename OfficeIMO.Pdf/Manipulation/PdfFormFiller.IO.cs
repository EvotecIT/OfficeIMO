namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static System.Threading.CancellationToken GetCancellationToken(PdfFormFillerOptions? options,
        System.Threading.CancellationToken cancellationToken = default) =>
        cancellationToken.CanBeCanceled ? cancellationToken : options?.CancellationToken ?? default;

    /// <summary>Combines an operation token with form options. The caller owns and disposes the returned source.</summary>
    internal static System.Threading.CancellationTokenSource? LinkCancellation(System.Threading.CancellationToken optionToken,
        ref System.Threading.CancellationToken cancellationToken) {
        if (optionToken.CanBeCanceled && cancellationToken.CanBeCanceled && optionToken != cancellationToken) {
            var linked = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(optionToken, cancellationToken);
            cancellationToken = linked.Token;
            return linked;
        }
        if (!cancellationToken.CanBeCanceled) cancellationToken = optionToken;
        return null;
    }

    private static byte[] ReadPath(string path, PdfFormFillerOptions? options = null, PdfLoadOptions? readOptions = null) =>
        PdfDocumentSource.FromPath(path, readOptions, GetCancellationToken(options)).Bytes;

    private static byte[] ReadStream(Stream stream, string paramName, PdfFormFillerOptions? options = null, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        return PdfDocumentSource.FromRemainingStream(stream, readOptions, GetCancellationToken(options)).Bytes;
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes, PdfFormFillerOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        System.Threading.CancellationToken cancellationToken = GetCancellationToken(options);
        const int chunkBytes = 64 * 1024;
        for (int offset = 0; offset < bytes.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkBytes, bytes.Length - offset);
            outputStream.Write(bytes, offset, count);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
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

        string fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }
}
