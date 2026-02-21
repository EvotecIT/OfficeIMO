using System.Globalization;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Shared input-size guard helpers for reader adapters.
/// </summary>
public static class ReaderInputLimits {
    /// <summary>
    /// Enforces <paramref name="maxBytes"/> against file length when available.
    /// </summary>
    public static void EnforceFileSize(string path, long? maxBytes) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!maxBytes.HasValue) return;

        try {
            var fi = new FileInfo(path);
            if (fi.Length > maxBytes.Value) {
                throw new IOException(
                    $"Input exceeds MaxInputBytes ({fi.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (IOException) {
            throw;
        } catch {
            // If file metadata cannot be read, do not block reads.
        }
    }

    /// <summary>
    /// Enforces <paramref name="maxBytes"/> against stream length when seekable.
    /// </summary>
    public static void EnforceSeekableStreamSize(Stream stream, long? maxBytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!maxBytes.HasValue) return;
        if (!stream.CanSeek) return;

        try {
            if (stream.Length > maxBytes.Value) {
                throw new IOException(
                    $"Input exceeds MaxInputBytes ({stream.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (NotSupportedException) {
            // ignore
        }
    }

    /// <summary>
    /// Ensures a seekable stream for parsers that require rewind/index operations.
    /// Non-seekable inputs are snapshotted into memory with <paramref name="maxInputBytes"/> enforcement.
    /// </summary>
    public static Stream EnsureSeekableReadStream(Stream stream, long? maxInputBytes, CancellationToken cancellationToken, out bool ownsStream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        if (stream.CanSeek) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            ownsStream = false;
            return stream;
        }

        var buffer = new MemoryStream();
        try {
            var chunk = new byte[64 * 1024];
            long totalBytes = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                var read = stream.Read(chunk, 0, chunk.Length);
                if (read <= 0) break;
                buffer.Write(chunk, 0, read);

                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        $"Input exceeds MaxInputBytes ({totalBytes.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                }
            }
        } catch {
            buffer.Dispose();
            throw;
        }

        buffer.Position = 0;
        ownsStream = true;
        return buffer;
    }
}
