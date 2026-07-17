using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class ContentLineDocumentIO {
    internal static byte[] Read(string filePath, ContentLineReaderOptions options,
        CancellationToken cancellationToken) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = File.OpenRead(filePath)) return Read(stream, options, cancellationToken);
    }

    internal static byte[] Read(Stream stream, ContentLineReaderOptions options,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return EmailByteReader.ReadAll(stream, options.MaxInputBytes,
            nameof(ContentLineReaderOptions.MaxInputBytes), cancellationToken);
    }

    internal static async Task<byte[]> ReadAsync(string filePath, ContentLineReaderOptions options,
        CancellationToken cancellationToken) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            return await ReadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        }
    }

    internal static Task<byte[]> ReadAsync(Stream stream, ContentLineReaderOptions options,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return EmailByteReader.ReadAllAsync(stream, options.MaxInputBytes,
            nameof(ContentLineReaderOptions.MaxInputBytes), cancellationToken);
    }

    internal static void Write(string filePath, byte[] bytes) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        OfficeFileCommit.WriteAllBytes(filePath, bytes);
    }

    internal static void Write(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
    }

    internal static Task WriteAsync(string filePath, byte[] bytes, CancellationToken cancellationToken) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        return OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken);
    }

    internal static Task WriteAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        return OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken);
    }
}
