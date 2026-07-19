using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Shared input-size guard helpers for reader adapters.
/// </summary>
public static class ReaderInputLimits {
    internal static MemoryStream CreateSnapshotStream(int initialCapacity = 0) {
        return new ReaderSnapshotStream(initialCapacity);
    }

    internal static bool IsSnapshotStream(Stream stream) {
        return stream is ReaderSnapshotStream;
    }

    /// <summary>
    /// Transfers the exact backing buffer of an internal Reader snapshot to a trusted format adapter.
    /// The snapshot must not be written after a successful transfer.
    /// </summary>
    internal static bool TryGetOwnedSnapshotBytes(Stream stream, out byte[] bytes) {
        if (stream is not ReaderSnapshotStream snapshot || snapshot.Length > int.MaxValue) {
            bytes = Array.Empty<byte>();
            return false;
        }

        int length = checked((int)snapshot.Length);
        if (snapshot.Capacity != length) {
            snapshot.Capacity = length;
        }

        bytes = snapshot.GetBuffer();
        return bytes.Length == length;
    }

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
    /// Creates a seekable snapshot for parsers that require rewind/index operations.
    /// Seekable inputs are read from the beginning and restored to their original position.
    /// Non-seekable inputs are read from their current forward position.
    /// </summary>
    public static Stream EnsureSeekableReadStream(Stream stream, long? maxInputBytes, CancellationToken cancellationToken, out bool ownsStream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        if (stream is ReaderSnapshotStream) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            stream.Position = 0;
            ownsStream = false;
            return stream;
        }

        bool restorePosition = stream.CanSeek;
        long originalPosition = 0;
        if (restorePosition) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            originalPosition = stream.Position;
            stream.Position = 0;
        }

        var buffer = new ReaderSnapshotStream(0);
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
        } finally {
            if (restorePosition) stream.Position = originalPosition;
        }

        buffer.Position = 0;
        ownsStream = true;
        return buffer;
    }

    /// <summary>
    /// Asynchronously creates a seekable stream snapshot for parsers that require rewind/index operations.
    /// Seekable inputs are read from the beginning and restored to their original position. Non-seekable inputs
    /// are read from their current forward position. The returned snapshot must be disposed by the caller.
    /// </summary>
    public static async Task<Stream> EnsureSeekableReadStreamAsync(
        Stream stream,
        long? maxInputBytes,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        cancellationToken.ThrowIfCancellationRequested();
        if (stream is ReaderSnapshotStream) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            stream.Position = 0;
            return stream;
        }

        bool restorePosition = stream.CanSeek;
        long originalPosition = 0;
        if (restorePosition) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            originalPosition = stream.Position;
            stream.Position = 0;
        }

        var buffer = new ReaderSnapshotStream(0);
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
        } finally {
            if (restorePosition) stream.Position = originalPosition;
        }

        buffer.Position = 0;
        return buffer;
    }

    private sealed class ReaderSnapshotStream : MemoryStream {
        internal ReaderSnapshotStream(int initialCapacity) : base(initialCapacity) {
        }
    }
}
