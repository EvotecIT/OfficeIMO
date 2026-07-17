namespace OfficeIMO.OneNote;

/// <summary>
/// Performs bounded physical-header inspection for OneNote files.
/// </summary>
public static class OneNoteFileProbe {
    /// <summary>Reads a OneNote header from a file path.</summary>
    public static OneNoteFileHeader ReadHeader(string path, OneNoteReaderOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));
        var effectiveOptions = PrepareOptions(options);
        var file = new FileInfo(path);
        if (!file.Exists) throw new FileNotFoundException("The OneNote file does not exist.", path);
        EnforceInputLength(file.Length, effectiveOptions);
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
            return ReadHeaderCore(stream, file.Length, effectiveOptions);
        }
    }

    /// <summary>Asynchronously reads a OneNote header from a file path.</summary>
    public static async Task<OneNoteFileHeader> ReadHeaderAsync(
        string path,
        OneNoteReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));
        var effectiveOptions = PrepareOptions(options);
        var file = new FileInfo(path);
        if (!file.Exists) throw new FileNotFoundException("The OneNote file does not exist.", path);
        EnforceInputLength(file.Length, effectiveOptions);
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 4096, true)) {
            return await ReadHeaderCoreAsync(stream, file.Length, effectiveOptions, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Reads a OneNote header from a caller-owned stream. Seekable streams are read from the
    /// beginning and restored to their original position. Non-seekable streams are consumed
    /// only through the physical header prefix.
    /// </summary>
    public static OneNoteFileHeader ReadHeader(Stream stream, OneNoteReaderOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        var effectiveOptions = PrepareOptions(options);
        long? length = TryGetLength(stream);
        if (length.HasValue) EnforceInputLength(length.Value, effectiveOptions);
        return ReadHeaderCore(stream, length, effectiveOptions);
    }

    /// <summary>
    /// Asynchronously reads a OneNote header from a caller-owned stream. Seekable streams are
    /// restored to their original position after the read.
    /// </summary>
    public static Task<OneNoteFileHeader> ReadHeaderAsync(
        Stream stream,
        OneNoteReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        var effectiveOptions = PrepareOptions(options);
        long? length = TryGetLength(stream);
        if (length.HasValue) EnforceInputLength(length.Value, effectiveOptions);
        return ReadHeaderCoreAsync(stream, length, effectiveOptions, cancellationToken);
    }

    private static OneNoteFileHeader ReadHeaderCore(Stream stream, long? length, OneNoteReaderOptions options) {
        bool restore = stream.CanSeek;
        long originalPosition = 0;
        if (restore) {
            originalPosition = stream.Position;
            stream.Position = 0;
        }

        try {
            byte[] prefix = ReadPrefix(stream, OneNoteFormatConstants.RevisionStoreHeaderLength);
            return OneNoteFileHeaderReader.Read(prefix, length, options);
        } finally {
            if (restore) stream.Position = originalPosition;
        }
    }

    private static async Task<OneNoteFileHeader> ReadHeaderCoreAsync(
        Stream stream,
        long? length,
        OneNoteReaderOptions options,
        CancellationToken cancellationToken) {
        bool restore = stream.CanSeek;
        long originalPosition = 0;
        if (restore) {
            originalPosition = stream.Position;
            stream.Position = 0;
        }

        try {
            byte[] prefix = await ReadPrefixAsync(stream, OneNoteFormatConstants.RevisionStoreHeaderLength, cancellationToken).ConfigureAwait(false);
            return OneNoteFileHeaderReader.Read(prefix, length, options);
        } finally {
            if (restore) stream.Position = originalPosition;
        }
    }

    private static byte[] ReadPrefix(Stream stream, int maxBytes) {
        var buffer = new byte[maxBytes];
        int total = 0;
        while (total < buffer.Length) {
            int read = stream.Read(buffer, total, buffer.Length - total);
            if (read <= 0) break;
            total += read;
        }
        if (total == buffer.Length) return buffer;
        var exact = new byte[total];
        Buffer.BlockCopy(buffer, 0, exact, 0, total);
        return exact;
    }

    private static async Task<byte[]> ReadPrefixAsync(Stream stream, int maxBytes, CancellationToken cancellationToken) {
        var buffer = new byte[maxBytes];
        int total = 0;
        while (total < buffer.Length) {
            int read = await stream.ReadAsync(buffer, total, buffer.Length - total, cancellationToken).ConfigureAwait(false);
            if (read <= 0) break;
            total += read;
        }
        if (total == buffer.Length) return buffer;
        var exact = new byte[total];
        Buffer.BlockCopy(buffer, 0, exact, 0, total);
        return exact;
    }

    private static OneNoteReaderOptions PrepareOptions(OneNoteReaderOptions? options) {
        var effective = options ?? new OneNoteReaderOptions();
        effective.Validate();
        return effective;
    }

    private static void EnforceInputLength(long length, OneNoteReaderOptions options) {
        if (options.MaxInputBytes.HasValue && length > options.MaxInputBytes.Value) {
            throw new IOException("OneNote input exceeds MaxInputBytes (" + length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " > " + options.MaxInputBytes.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + ").");
        }
    }

    private static long? TryGetLength(Stream stream) {
        if (!stream.CanSeek) return null;
        try {
            return stream.Length;
        } catch (NotSupportedException) {
            return null;
        }
    }
}
