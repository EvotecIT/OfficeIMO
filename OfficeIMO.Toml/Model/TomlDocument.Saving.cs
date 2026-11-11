namespace OfficeIMO.Toml;

public sealed partial class TomlDocument {
    private const int EncodeChunk = 8192;
    private static readonly byte[] Utf8Bom = { 0xEF, 0xBB, 0xBF };

    /// <summary>Saves the document to the file it was loaded from or created for.</summary>
    /// <exception cref="InvalidOperationException">The document was parsed from text or a stream and has no file path.</exception>
    public void Save() {
        if (FilePath is null) throw new InvalidOperationException("The document has no file path; save it with Save(path) or Save(stream).");
        Save(FilePath);
    }

    /// <summary>
    /// Saves the document to a file as UTF-8, with a byte-order mark when the input had one. An unchanged document
    /// reproduces the loaded bytes exactly.
    /// </summary>
    /// <param name="path">The file to create or overwrite.</param>
    public void Save(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        List<TomlTextSegment> segments = TomlRendering.Collect(this);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize);
        Write(stream, segments);
    }

    /// <summary>Saves the document to a caller-owned stream as UTF-8, with a byte-order mark when the input had one. The stream is not disposed.</summary>
    /// <param name="stream">A writable stream.</param>
    public void Save(Stream stream) {
        RequireWritable(stream);
        Write(stream, TomlRendering.Collect(this));
    }

    /// <summary>Saves the document to the file it was loaded from or created for, asynchronously.</summary>
    /// <param name="cancellationToken">Observed between writes.</param>
    /// <exception cref="InvalidOperationException">The document has no file path.</exception>
    public Task SaveAsync(CancellationToken cancellationToken = default) {
        if (FilePath is null) throw new InvalidOperationException("The document has no file path; save it with SaveAsync(path) or SaveAsync(stream).");
        return SaveAsync(FilePath, cancellationToken);
    }

    /// <summary>Saves the document to a file asynchronously; see <see cref="Save(string)"/>.</summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        List<TomlTextSegment> segments = TomlRendering.Collect(this);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.Asynchronous);
        await WriteAsync(stream, segments, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the document to a caller-owned stream asynchronously; see <see cref="Save(Stream)"/>.</summary>
    /// <param name="stream">A writable stream. It is not disposed.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public Task SaveAsync(Stream stream, CancellationToken cancellationToken = default) {
        RequireWritable(stream);
        return WriteAsync(stream, TomlRendering.Collect(this), cancellationToken);
    }

    /// <summary>
    /// Returns the document as TOML text. An unchanged document returns its source text: the string passed to
    /// <see cref="Parse"/>, or the decoded file or stream without its byte-order mark.
    /// </summary>
    public override string ToString() {
        List<TomlTextSegment> segments = TomlRendering.Collect(this);
        if (segments.Count == 1 && segments[0].Start == 0 && segments[0].Length == segments[0].Text.Length) return segments[0].Text;
        var text = new StringBuilder(_source.Length + 64);
        for (int i = 0; i < segments.Count; i++) text.Append(segments[i].Text, segments[i].Start, segments[i].Length);
        return text.ToString();
    }

    private void Write(Stream stream, List<TomlTextSegment> segments) {
        Encoder encoder = StrictUtf8.GetEncoder();
        char[] chars = ArrayPool<char>.Shared.Rent(EncodeChunk);
        byte[] bytes = ArrayPool<byte>.Shared.Rent(EncodeChunk * 3 + 8);
        try {
            if (_byteOrderMark) stream.Write(Utf8Bom, 0, Utf8Bom.Length);
            for (int s = 0; s < segments.Count; s++) {
                TomlTextSegment segment = segments[s];
                for (int offset = 0; offset < segment.Length; offset += EncodeChunk) {
                    int written = Encode(encoder, segment, offset, chars, bytes);
                    stream.Write(bytes, 0, written);
                }
            }
            int tail = encoder.GetBytes(chars, 0, 0, bytes, 0, flush: true);
            if (tail > 0) stream.Write(bytes, 0, tail);
        } finally {
            ArrayPool<char>.Shared.Return(chars);
            ArrayPool<byte>.Shared.Return(bytes);
        }
    }

    private async Task WriteAsync(Stream stream, List<TomlTextSegment> segments, CancellationToken cancellationToken) {
        Encoder encoder = StrictUtf8.GetEncoder();
        char[] chars = ArrayPool<char>.Shared.Rent(EncodeChunk);
        byte[] bytes = ArrayPool<byte>.Shared.Rent(EncodeChunk * 3 + 8);
        try {
            if (_byteOrderMark) await stream.WriteAsync(Utf8Bom, 0, Utf8Bom.Length, cancellationToken).ConfigureAwait(false);
            for (int s = 0; s < segments.Count; s++) {
                TomlTextSegment segment = segments[s];
                for (int offset = 0; offset < segment.Length; offset += EncodeChunk) {
                    int written = Encode(encoder, segment, offset, chars, bytes);
                    await stream.WriteAsync(bytes, 0, written, cancellationToken).ConfigureAwait(false);
                }
            }
            int tail = encoder.GetBytes(chars, 0, 0, bytes, 0, flush: true);
            if (tail > 0) await stream.WriteAsync(bytes, 0, tail, cancellationToken).ConfigureAwait(false);
        } finally {
            ArrayPool<char>.Shared.Return(chars);
            ArrayPool<byte>.Shared.Return(bytes);
        }
    }

    private static int Encode(Encoder encoder, TomlTextSegment segment, int offset, char[] chars, byte[] bytes) {
        int count = Math.Min(EncodeChunk, segment.Length - offset);
        segment.Text.CopyTo(segment.Start + offset, chars, 0, count);
        return encoder.GetBytes(chars, 0, count, bytes, 0, flush: false);
    }

    private static void RequireWritable(Stream stream) {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
    }
}
