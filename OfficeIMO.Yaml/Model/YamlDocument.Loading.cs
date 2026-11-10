namespace OfficeIMO.Yaml;

public sealed partial class YamlDocument {
    private const int FileBufferSize = 81920;

    /// <summary>
    /// Loads a YAML file. The encoding (UTF-8, UTF-16 or UTF-32, either byte order) is detected from a byte-order mark or
    /// the first bytes as the YAML specification requires, and recorded so <see cref="Save(string, YamlSaveOptions?)"/> reproduces it.
    /// </summary>
    /// <param name="path">The file to read.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="YamlException">The input is malformed, not valid in its encoding, or exceeds a limit.</exception>
    /// <param name="cancellationToken">Observed while reading and, at bounded intervals, while parsing.</param>
    public static YamlDocument Load(string path, YamlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.SequentialScan);
        YamlDocument document = Load(stream, options, cancellationToken);
        document._path = path;
        return document;
    }

    /// <summary>
    /// Loads YAML from a caller-owned stream, reading from its current position to the end. The encoding is detected as
    /// for <see cref="Load(string, YamlLoadOptions?, CancellationToken)"/>. The stream is not disposed.
    /// </summary>
    /// <param name="stream">A readable stream.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="YamlException">The input is malformed, not valid in its encoding, or exceeds a limit.</exception>
    /// <param name="cancellationToken">Observed between reads and, at bounded intervals, while parsing.</param>
    public static YamlDocument Load(Stream stream, YamlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        RequireReadable(stream);
        options ??= YamlLoadOptions.Default;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(InitialCapacity(stream));
        try {
            int count = 0;
            while (true) {
                if (count == buffer.Length) buffer = Grow(buffer, count);
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(buffer, count, buffer.Length - count);
                if (read <= 0) break;
                count += read;
                CheckInputSize(count, options);
            }
            return FromBytes(buffer, count, options, cancellationToken);
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    /// <summary>Loads a YAML file asynchronously; see <see cref="Load(string, YamlLoadOptions?, CancellationToken)"/>.</summary>
    /// <param name="path">The file to read.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed while reading and, at bounded intervals, while parsing.</param>
    /// <exception cref="YamlException">The input is malformed, not valid in its encoding, or exceeds a limit.</exception>
    public static async Task<YamlDocument> LoadAsync(string path, YamlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.Asynchronous | FileOptions.SequentialScan);
        YamlDocument document = await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        document._path = path;
        return document;
    }

    /// <summary>Loads YAML from a caller-owned stream asynchronously; see <see cref="Load(Stream, YamlLoadOptions?, CancellationToken)"/>.</summary>
    /// <param name="stream">A readable stream. It is not disposed.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed while reading and, at bounded intervals, while parsing.</param>
    /// <exception cref="YamlException">The input is malformed, not valid in its encoding, or exceeds a limit.</exception>
    public static async Task<YamlDocument> LoadAsync(Stream stream, YamlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        RequireReadable(stream);
        options ??= YamlLoadOptions.Default;
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
            return FromBytes(buffer, count, options, cancellationToken);
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    /// <summary>
    /// Parses YAML text. The text is retained as is, so <see cref="ToString()"/> returns it unchanged and
    /// <see cref="Save(Stream, YamlSaveOptions?)"/> writes it as UTF-8 without adding a byte-order mark.
    /// </summary>
    /// <param name="yaml">The YAML text.</param>
    /// <param name="options">Load settings, or <see langword="null"/> for the defaults.</param>
    /// <exception cref="YamlException">The input is malformed or exceeds a limit.</exception>
    /// <param name="cancellationToken">Checked at bounded intervals while parsing.</param>
    public static YamlDocument Parse(string yaml, YamlLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        return Build(yaml, YamlTextEncoding.Utf8, byteOrderMark: false, options ?? YamlLoadOptions.Default, cancellationToken);
    }

    private static YamlDocument FromBytes(byte[] buffer, int count, YamlLoadOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string text = YamlEncodingDetector.Decode(buffer, count, out YamlTextEncoding encoding, out bool byteOrderMark);
        return Build(text, encoding, byteOrderMark, options, cancellationToken);
    }

    private static YamlDocument Build(string text, YamlTextEncoding encoding, bool byteOrderMark, YamlLoadOptions options, CancellationToken cancellationToken) {
        var document = new YamlDocument(text, encoding, byteOrderMark, options.DuplicateKeys, options.Strict);
        try {
            using var reader = new YamlReader(text.AsMemory(), options.ReaderOptions, cancellationToken);
            new YamlTreeBuilder(document, options).Run(reader);
        } catch (YamlException exception) when (exception.AttachSource(text)) {
            throw;
        }
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
        if (count >= int.MaxValue - 64) throw new YamlLimitExceededException("MaxInputLength", int.MaxValue - 64, new YamlPosition(count, 0, 0));
        byte[] larger = ArrayPool<byte>.Shared.Rent((int)Math.Min((long)buffer.Length * 2, int.MaxValue - 64));
        Buffer.BlockCopy(buffer, 0, larger, 0, count);
        ArrayPool<byte>.Shared.Return(buffer);
        return larger;
    }

    private static void CheckInputSize(int count, YamlLoadOptions options) {
        if (options.MaxInputBytes is long max && count > max) {
            throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxInputBytes), max, new YamlPosition(max, 0, 0));
        }
    }
}
