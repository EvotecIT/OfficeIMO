using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

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
    /// Enforces <paramref name="maxBytes"/> against the unread portion of a seekable stream.
    /// </summary>
    public static void EnforceSeekableStreamRemainingSize(Stream stream, long? maxBytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!maxBytes.HasValue) return;
        if (!stream.CanSeek) return;

        try {
            long remainingBytes = Math.Max(0L, stream.Length - stream.Position);
            if (remainingBytes > maxBytes.Value) {
                throw new IOException(
                    $"Input exceeds MaxInputBytes ({remainingBytes.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
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
            EnforceSeekableStreamRemainingSize(stream, maxInputBytes);
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

    /// <summary>
    /// Asynchronously ensures a seekable stream for parsers that require rewind/index operations.
    /// The original stream is returned when it is seekable. Otherwise a bounded memory snapshot is returned
    /// and must be disposed by the caller.
    /// </summary>
    public static async Task<Stream> EnsureSeekableReadStreamAsync(
        Stream stream,
        long? maxInputBytes,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            EnforceSeekableStreamRemainingSize(stream, maxInputBytes);
            return stream;
        }

        var buffer = new MemoryStream();
        try {
            var chunk = new byte[64 * 1024];
            long totalBytes = 0;
            while (true) {
                int read = await stream.ReadAsync(chunk, 0, chunk.Length, cancellationToken).ConfigureAwait(false);
                if (read <= 0) break;

                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        $"Input exceeds MaxInputBytes ({totalBytes.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                }

                await buffer.WriteAsync(chunk, 0, read, cancellationToken).ConfigureAwait(false);
            }
        } catch {
            buffer.Dispose();
            throw;
        }

        buffer.Position = 0;
        return buffer;
    }
}
