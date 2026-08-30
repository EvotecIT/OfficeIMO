using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Shared input-size guard helpers for reader adapters.
/// </summary>
public static class ReaderInputLimits {
    private const long MaximumInMemorySnapshotBytes = 64L * 1024 * 1024;
    private const uint OwnerDirectoryMode = 0x1C0; // 0700

    internal static MemoryStream CreateSnapshotStream(int initialCapacity = 0) {
        return new ReaderSnapshotStream(initialCapacity);
    }

    internal static bool IsSnapshotStream(Stream stream) {
        return stream is ReaderSnapshotStream
            || stream is ReaderSnapshotFileStream;
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
    public static Stream EnsureSeekableReadStream(Stream stream, long? maxInputBytes, CancellationToken cancellationToken, out bool ownsStream) =>
        EnsureSeekableReadStream(stream, maxInputBytes, inputLimitProbe: null, cancellationToken, out ownsStream);

    internal static Stream EnsureSeekableReadStream(Stream stream, long? maxInputBytes, ReaderInputLimitProbe? inputLimitProbe,
        CancellationToken cancellationToken, out bool ownsStream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        if (IsSnapshotStream(stream)) {
            long snapshotOriginalPosition = stream.Position;
            try {
                stream.Position = 0;
                maxInputBytes = ResolveProbedMaxInputBytes(stream, maxInputBytes, inputLimitProbe, cancellationToken);
                EnforceSeekableStreamSize(stream, maxInputBytes);
                stream.Position = 0;
            } catch {
                stream.Position = snapshotOriginalPosition;
                throw;
            }
            ownsStream = false;
            return stream;
        }

        bool restorePosition = stream.CanSeek;
        long originalPosition = 0;
        if (restorePosition) {
            originalPosition = stream.Position;
            try {
                stream.Position = 0;
                maxInputBytes = ResolveProbedMaxInputBytes(stream, maxInputBytes, inputLimitProbe, cancellationToken);
                EnforceSeekableStreamSize(stream, maxInputBytes);
                stream.Position = 0;
            } catch {
                stream.Position = originalPosition;
                throw;
            }
        }

        Stream buffer = CreateBoundedSnapshotBuffer(maxInputBytes);
        try {
            var chunk = new byte[64 * 1024];
            byte[]? prefix = inputLimitProbe == null ? null : new byte[inputLimitProbe.PrefixLength];
            int prefixLength = 0;
            bool probeResolved = inputLimitProbe == null;
            long totalBytes = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                var read = stream.Read(chunk, 0, chunk.Length);
                if (read <= 0) break;

                if (!probeResolved && prefix != null) {
                    int copy = Math.Min(read, prefix.Length - prefixLength);
                    Buffer.BlockCopy(chunk, 0, prefix, prefixLength, copy);
                    prefixLength += copy;
                    if (prefixLength == prefix.Length) {
                        maxInputBytes = CombineMaxInputBytes(maxInputBytes,
                            inputLimitProbe!.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, prefixLength)));
                        probeResolved = true;
                    }
                }

                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        $"Input exceeds MaxInputBytes ({totalBytes.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                }
                buffer.Write(chunk, 0, read);
            }
            if (!probeResolved && prefix != null) {
                maxInputBytes = CombineMaxInputBytes(maxInputBytes,
                    inputLimitProbe!.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, prefixLength)));
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
        CancellationToken cancellationToken = default) =>
        await EnsureSeekableReadStreamAsync(stream, maxInputBytes, inputLimitProbe: null, cancellationToken).ConfigureAwait(false);

    internal static async Task<Stream> EnsureSeekableReadStreamAsync(
        Stream stream,
        long? maxInputBytes,
        ReaderInputLimitProbe? inputLimitProbe,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        cancellationToken.ThrowIfCancellationRequested();
        if (IsSnapshotStream(stream)) {
            long snapshotOriginalPosition = stream.Position;
            try {
                stream.Position = 0;
                maxInputBytes = await ResolveProbedMaxInputBytesAsync(stream, maxInputBytes, inputLimitProbe, cancellationToken).ConfigureAwait(false);
                EnforceSeekableStreamSize(stream, maxInputBytes);
                stream.Position = 0;
            } catch {
                stream.Position = snapshotOriginalPosition;
                throw;
            }
            return stream;
        }

        bool restorePosition = stream.CanSeek;
        long originalPosition = 0;
        if (restorePosition) {
            originalPosition = stream.Position;
            try {
                stream.Position = 0;
                maxInputBytes = await ResolveProbedMaxInputBytesAsync(stream, maxInputBytes, inputLimitProbe, cancellationToken).ConfigureAwait(false);
                EnforceSeekableStreamSize(stream, maxInputBytes);
                stream.Position = 0;
            } catch {
                stream.Position = originalPosition;
                throw;
            }
        }

        Stream buffer = CreateBoundedSnapshotBuffer(maxInputBytes);
        try {
            var chunk = new byte[64 * 1024];
            byte[]? prefix = inputLimitProbe == null ? null : new byte[inputLimitProbe.PrefixLength];
            int prefixLength = 0;
            bool probeResolved = inputLimitProbe == null;
            long totalBytes = 0;
            while (true) {
                int read = await stream.ReadAsync(chunk, 0, chunk.Length, cancellationToken).ConfigureAwait(false);
                if (read <= 0) break;

                if (!probeResolved && prefix != null) {
                    int copy = Math.Min(read, prefix.Length - prefixLength);
                    Buffer.BlockCopy(chunk, 0, prefix, prefixLength, copy);
                    prefixLength += copy;
                    if (prefixLength == prefix.Length) {
                        maxInputBytes = CombineMaxInputBytes(maxInputBytes,
                            inputLimitProbe!.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, prefixLength)));
                        probeResolved = true;
                    }
                }

                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        $"Input exceeds MaxInputBytes ({totalBytes.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                }

                await buffer.WriteAsync(chunk, 0, read, cancellationToken).ConfigureAwait(false);
            }
            if (!probeResolved && prefix != null) {
                maxInputBytes = CombineMaxInputBytes(maxInputBytes,
                    inputLimitProbe!.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, prefixLength)));
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
        return buffer;
    }

    private static long? ResolveProbedMaxInputBytes(Stream stream, long? maxInputBytes,
        ReaderInputLimitProbe? inputLimitProbe, CancellationToken cancellationToken) {
        if (inputLimitProbe == null) return maxInputBytes;
        var prefix = new byte[inputLimitProbe.PrefixLength];
        int total = 0;
        while (total < prefix.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(prefix, total, prefix.Length - total);
            if (read <= 0) break;
            total += read;
        }
        stream.Position = 0;
        return CombineMaxInputBytes(maxInputBytes,
            inputLimitProbe.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, total)));
    }

    private static async Task<long?> ResolveProbedMaxInputBytesAsync(Stream stream, long? maxInputBytes,
        ReaderInputLimitProbe? inputLimitProbe, CancellationToken cancellationToken) {
        if (inputLimitProbe == null) return maxInputBytes;
        var prefix = new byte[inputLimitProbe.PrefixLength];
        int total = 0;
        while (total < prefix.Length) {
            int read = await stream.ReadAsync(prefix, total, prefix.Length - total, cancellationToken).ConfigureAwait(false);
            if (read <= 0) break;
            total += read;
        }
        stream.Position = 0;
        return CombineMaxInputBytes(maxInputBytes,
            inputLimitProbe.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, total)));
    }

    private static long? CombineMaxInputBytes(long? configured, long? probed) {
        if (probed.HasValue && probed.Value < 1) throw new InvalidOperationException("An input-limit prefix resolver returned a value below 1.");
        return configured.HasValue && probed.HasValue ? Math.Min(configured.Value, probed.Value) : configured ?? probed;
    }

    private static Stream CreateBoundedSnapshotBuffer(long? maxInputBytes) {
        if (maxInputBytes.HasValue
            && maxInputBytes.Value <= MaximumInMemorySnapshotBytes) {
            return new ReaderSnapshotStream(0);
        }

        return CreatePrivateSnapshotFileStream();
    }

    private static Stream CreatePrivateSnapshotFileStream() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Directory.CreateDirectory(directory);
        } else if (CreateDirectoryUnix(directory, OwnerDirectoryMode) != 0) {
            throw new IOException(
                "Unable to create the private Reader snapshot directory (OS error " +
                Marshal.GetLastWin32Error().ToString(CultureInfo.InvariantCulture) + ").");
        }
        try {
            string path = Path.Combine(directory, "snapshot.tmp");
            return new ReaderSnapshotFileStream(path, directory);
        } catch {
            TryDeleteSnapshotDirectory(directory);
            throw;
        }
    }

    private static void TryDeleteSnapshotDirectory(string directory) {
        try {
            Directory.Delete(directory, recursive: true);
        } catch (IOException) {
            // DeleteOnClose remains the primary cleanup; directory removal is best effort.
        } catch (UnauthorizedAccessException) {
            // DeleteOnClose remains the primary cleanup; directory removal is best effort.
        }
    }

    private sealed class ReaderSnapshotStream : MemoryStream {
        internal ReaderSnapshotStream(int initialCapacity) : base(initialCapacity) {
        }
    }

    private sealed class ReaderSnapshotFileStream : FileStream {
        private readonly string _directory;

        internal ReaderSnapshotFileStream(string path, string directory) : base(path,
            FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            64 * 1024,
            FileOptions.DeleteOnClose | FileOptions.SequentialScan) {
            _directory = directory;
        }

        protected override void Dispose(bool disposing) {
            try {
                base.Dispose(disposing);
            } finally {
                if (disposing) TryDeleteSnapshotDirectory(_directory);
            }
        }

#if NET6_0_OR_GREATER
        public override async ValueTask DisposeAsync() {
            try {
                await base.DisposeAsync().ConfigureAwait(false);
            } finally {
                TryDeleteSnapshotDirectory(_directory);
            }
        }
#endif
    }

    [DllImport("libc", EntryPoint = "mkdir", SetLastError = true)]
    private static extern int CreateDirectoryUnix(string path, uint mode);
}
