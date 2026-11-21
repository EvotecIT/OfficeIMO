namespace OfficeIMO.Toml;

public sealed partial class TomlDocument {
    private const int FileBufferSize = 81920;
    private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    /// <summary>
    /// Loads a TOML file. TOML is UTF-8; a UTF-8 byte-order mark is accepted and recorded so <see cref="Save()"/> writes
    /// it back. The path is remembered for <see cref="Save()"/>.
    /// </summary>
    /// <param name="path">The file to read.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="TomlException">The input is not valid UTF-8 or TOML, or exceeds a limit.</exception>
    public static TomlDocument Load(string path, TomlLoadOptions? options = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.SequentialScan);
        TomlDocument document = Load(stream, options);
        document.FilePath = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Loads TOML from a caller-owned stream, reading from its current position to the end. The stream is not disposed.</summary>
    /// <param name="stream">A readable stream.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="TomlException">The input is not valid UTF-8 or TOML, or exceeds a limit.</exception>
    public static TomlDocument Load(Stream stream, TomlLoadOptions? options = null) {
        options ??= TomlLoadOptions.Default;
        string text = ReadText(stream, options, out bool bom);
        return Build(text, bom, options, CancellationToken.None);
    }

    /// <summary>Reads a stream to its end and decodes it as <see cref="Load(Stream, TomlLoadOptions?)"/> does, without parsing it.</summary>
    internal static string ReadText(Stream stream, TomlLoadOptions options, out bool byteOrderMark) {
        RequireReadable(stream);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(InitialCapacity(stream));
        try {
            int count = 0;
            while (true) {
                if (count == buffer.Length) buffer = Grow(buffer, count);
                int read = stream.Read(buffer, count, buffer.Length - count);
                if (read <= 0) break;
                count += read;
                CheckInputSize(count, options);
            }
            return Decode(buffer, count, out byteOrderMark);
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    /// <summary>Loads a TOML file asynchronously; see <see cref="Load(string, TomlLoadOptions?)"/>.</summary>
    /// <param name="path">The file to read.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed while reading and, at bounded intervals, while parsing.</param>
    /// <exception cref="TomlException">The input is not valid UTF-8 or TOML, or exceeds a limit.</exception>
    public static async Task<TomlDocument> LoadAsync(string path, TomlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.Asynchronous | FileOptions.SequentialScan);
        TomlDocument document = await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        document.FilePath = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Loads TOML from a caller-owned stream asynchronously; see <see cref="Load(Stream, TomlLoadOptions?)"/>.</summary>
    /// <param name="stream">A readable stream. It is not disposed.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed while reading and, at bounded intervals, while parsing.</param>
    /// <exception cref="TomlException">The input is not valid UTF-8 or TOML, or exceeds a limit.</exception>
    public static async Task<TomlDocument> LoadAsync(Stream stream, TomlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        RequireReadable(stream);
        options ??= TomlLoadOptions.Default;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(InitialCapacity(stream));
        try {
            int count = 0;
            while (true) {
                if (count == buffer.Length) buffer = Grow(buffer, count);
                int read = await stream.ReadAsync(buffer, count, buffer.Length - count, cancellationToken).ConfigureAwait(false);
                if (read <= 0) break;
                count += read;
                CheckInputSize(count, options);
            }
            return FromBytes(buffer, count, options, null, cancellationToken);
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    /// <summary>
    /// Parses TOML text. The text is retained as is, so <see cref="ToString"/> returns it unchanged and
    /// <see cref="Save(Stream)"/> writes it as UTF-8 without a byte-order mark.
    /// </summary>
    /// <param name="toml">The TOML text.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="TomlException">The input is not valid TOML or exceeds a limit.</exception>
    public static TomlDocument Parse(string toml, TomlLoadOptions? options = null) {
        if (toml is null) throw new ArgumentNullException(nameof(toml));
        options ??= TomlLoadOptions.Default;
        if (options.MaxInputBytes is long max && toml.Length > max) {
            throw new TomlLimitExceededException(nameof(TomlLoadOptions.MaxInputBytes), max, default);
        }
        return Build(toml, byteOrderMark: false, options, CancellationToken.None);
    }

    private static TomlDocument FromBytes(byte[] buffer, int count, TomlLoadOptions options, string? path, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string text = Decode(buffer, count, out bool bom);
        return Build(text, bom, options, cancellationToken);
    }

    private static string Decode(byte[] buffer, int count, out bool byteOrderMark) {
        byteOrderMark = count >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF;
        int offset = byteOrderMark ? 3 : 0;
        try {
            return StrictUtf8.GetString(buffer, offset, count - offset);
        } catch (DecoderFallbackException ex) {
            // DecoderFallbackException.Index is not comparable across runtimes, so locate the first invalid byte with
            // the same rules the stream reader applies.
            int index = FirstInvalidUtf8(buffer, offset, count);
            throw new TomlException(TomlErrorCodes.InvalidEncoding, "The input is not valid UTF-8.", BytePosition(buffer, offset, index), ex);
        }
    }

    /// <summary>Returns the index of the first byte that starts an invalid or truncated UTF-8 sequence.</summary>
    private static int FirstInvalidUtf8(byte[] buffer, int start, int end) {
        int i = start;
        while (i < end) {
            byte lead = buffer[i];
            if (lead < 0x80) {
                i++;
                continue;
            }
            int need;
            byte low = 0x80, high = 0xBF;
            if (lead >= 0xC2 && lead <= 0xDF) {
                need = 1;
            } else if (lead >= 0xE0 && lead <= 0xEF) {
                need = 2;
                if (lead == 0xE0) low = 0xA0;
                else if (lead == 0xED) high = 0x9F;
            } else if (lead >= 0xF0 && lead <= 0xF4) {
                need = 3;
                if (lead == 0xF0) low = 0x90;
                else if (lead == 0xF4) high = 0x8F;
            } else {
                return i;
            }
            if (i + need >= end) return i;
            for (int j = 1; j <= need; j++) {
                byte next = buffer[i + j];
                if (j == 1 ? next < low || next > high : (next & 0xC0) != 0x80) return i;
            }
            i += need + 1;
        }
        return end;
    }

    /// <summary>Finds the line and column of an invalid byte, counting the UTF-16 code units decoded before it.</summary>
    private static TomlPosition BytePosition(byte[] buffer, int start, int index) {
        int line = 1, column = 1, offset = 0;
        for (int i = start; i < index && i < buffer.Length; i++) {
            byte b = buffer[i];
            if ((b & 0xC0) == 0x80) continue; // continuation byte
            int units = b >= 0xF0 ? 2 : 1; // four-byte sequences decode to a surrogate pair
            offset += units;
            if (b == (byte)'\n') {
                line++;
                column = 1;
            } else {
                column += units;
            }
        }
        return new TomlPosition(offset, line, column);
    }

    private static TomlDocument Build(string text, bool byteOrderMark, TomlLoadOptions options, CancellationToken cancellationToken) {
        var document = new TomlDocument(text, byteOrderMark, filePath: null);
        var reader = new TomlReader(text.AsMemory(), options.ReaderOptions, cancellationToken);
        new TomlTreeBuilder(document, options).Run(reader);
        return document;
    }

    private static void RequireReadable(Stream stream) {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
    }

    private static int InitialCapacity(Stream stream) {
        if (!stream.CanSeek) return FileBufferSize;
        long remaining = stream.Length - stream.Position;
        return remaining <= 0 ? 256 : remaining >= int.MaxValue - 64 ? FileBufferSize : (int)remaining + 1;
    }

    private static byte[] Grow(byte[] buffer, int count) {
        if (count >= int.MaxValue - 64) throw new TomlLimitExceededException("MaxInputLength", int.MaxValue - 64, new TomlPosition(count, 0, 0));
        byte[] larger = ArrayPool<byte>.Shared.Rent((int)Math.Min((long)buffer.Length * 2, int.MaxValue - 64));
        Buffer.BlockCopy(buffer, 0, larger, 0, count);
        ArrayPool<byte>.Shared.Return(buffer);
        return larger;
    }

    private static void CheckInputSize(int count, TomlLoadOptions options) {
        if (options.MaxInputBytes is long max && count > max) {
            throw new TomlLimitExceededException(nameof(TomlLoadOptions.MaxInputBytes), max, new TomlPosition(max, 0, 0));
        }
    }
}
