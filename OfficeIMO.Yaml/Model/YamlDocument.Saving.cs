namespace OfficeIMO.Yaml;

/// <summary>A run of output text: a range of a string, usually of the retained source.</summary>
internal readonly struct YamlTextSegment {
    public YamlTextSegment(string text, int start, int length) {
        Text = text;
        Start = start;
        Length = length;
    }

    public string Text { get; }
    public int Start { get; }
    public int Length { get; }
}

public sealed partial class YamlDocument {
    private const int EncodeChunk = 8192;

    /// <summary>
    /// Saves the document back to the file it was loaded from, in the encoding it was loaded with. Untouched text is
    /// copied from the source; only edited nodes are written anew.
    /// </summary>
    /// <exception cref="InvalidOperationException">The document was not loaded from a file path.</exception>
    public void Save() {
        if (_path is null) throw new InvalidOperationException("The document was not loaded from a file; use Save(path) or Save(stream).");
        Save(_path);
    }

    /// <summary>
    /// Saves the document to a file in the encoding it was loaded with, writing its byte-order mark when the input had one.
    /// An unchanged document reproduces the loaded bytes exactly; an edited one changes only the edited nodes' text.
    /// </summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    public void Save(string path, YamlSaveOptions? options = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        List<YamlTextSegment> segments = CollectSegments(options);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize);
        Write(stream, segments);
    }

    /// <summary>
    /// Saves the document to a caller-owned stream in the encoding it was loaded with, including its byte-order mark.
    /// The stream is not disposed.
    /// </summary>
    /// <param name="stream">A writable stream.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    public void Save(Stream stream, YamlSaveOptions? options = null) {
        RequireWritable(stream);
        Write(stream, CollectSegments(options));
    }

    /// <summary>Saves the document back to the file it was loaded from asynchronously; see <see cref="Save()"/>.</summary>
    /// <param name="cancellationToken">Observed between writes.</param>
    public Task SaveAsync(CancellationToken cancellationToken = default) {
        if (_path is null) throw new InvalidOperationException("The document was not loaded from a file; use SaveAsync(path) or SaveAsync(stream).");
        return SaveAsync(_path, null, cancellationToken);
    }

    /// <summary>Saves the document to a file asynchronously; see <see cref="Save(string, YamlSaveOptions?)"/>.</summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public async Task SaveAsync(string path, YamlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        List<YamlTextSegment> segments = CollectSegments(options);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.Asynchronous);
        await WriteAsync(stream, segments, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the document to a caller-owned stream asynchronously; see <see cref="Save(Stream, YamlSaveOptions?)"/>.</summary>
    /// <param name="stream">A writable stream. It is not disposed.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public Task SaveAsync(Stream stream, YamlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        RequireWritable(stream);
        return WriteAsync(stream, CollectSegments(options), cancellationToken);
    }

    /// <summary>
    /// Returns the document as YAML text, as <see cref="Save(Stream, YamlSaveOptions?)"/> would write it. An unchanged
    /// document returns its source text: the string passed to <see cref="Parse"/>, or the decoded file or stream without
    /// its byte-order mark.
    /// </summary>
    public override string ToString() => ToString(null);

    /// <summary>Returns the document as YAML text with save settings; see <see cref="ToString()"/>.</summary>
    public string ToString(YamlSaveOptions? options) {
        List<YamlTextSegment> segments = CollectSegments(options);
        if (segments.Count == 1 && segments[0].Start == 0 && segments[0].Length == segments[0].Text.Length) return segments[0].Text;
        int length = 0;
        for (int i = 0; i < segments.Count; i++) length += segments[i].Length;
        var text = new StringBuilder(length);
        for (int i = 0; i < segments.Count; i++) text.Append(segments[i].Text, segments[i].Start, segments[i].Length);
        return text.ToString();
    }

    /// <summary>
    /// Lists the output text: source ranges for everything clean, and emitted text for edits. An unchanged document is
    /// one range covering the whole source; a reformat writes everything through the emitter.
    /// </summary>
    internal List<YamlTextSegment> CollectSegments(YamlSaveOptions? options) {
        if (options?.Reformat == true) return Reformat(options.WriterOptions ?? WriterOptions);
        if (!_partsDirty) {
            var whole = new List<YamlTextSegment>(1);
            if (_source.Length > 0) whole.Add(new YamlTextSegment(_source, 0, _source.Length));
            return whole;
        }
        return new YamlSaveWalk(this).Run();
    }

    private List<YamlTextSegment> Reformat(YamlWriterOptions options) {
        var text = new StringWriter(CultureInfo.InvariantCulture);
        var emitter = new YamlEmitter(text, options);
        for (int i = 0; i < _parts.Count; i++) YamlModelWriter.WritePart(emitter, _parts[i], marker: _parts[i].HasStartMarker || i > 0);
        if (_parts.Count > 0) emitter.Finish();
        string output = text.ToString();
        return new List<YamlTextSegment>(1) { new YamlTextSegment(output, 0, output.Length) };
    }

    private void Write(Stream stream, List<YamlTextSegment> segments) {
        Encoder encoder = YamlEncodingDetector.CreateEncoding(_encoding).GetEncoder();
        char[] chars = ArrayPool<char>.Shared.Rent(EncodeChunk);
        byte[] bytes = ArrayPool<byte>.Shared.Rent(EncodeChunk * 4 + 8);
        try {
            if (_byteOrderMark) {
                byte[] mark = YamlEncodingDetector.GetByteOrderMark(_encoding);
                stream.Write(mark, 0, mark.Length);
            }
            for (int s = 0; s < segments.Count; s++) {
                YamlTextSegment segment = segments[s];
                for (int offset = 0; offset < segment.Length; offset += EncodeChunk) {
                    int written = Encode(encoder, segment, offset, chars, bytes, flush: false);
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

    private async Task WriteAsync(Stream stream, List<YamlTextSegment> segments, CancellationToken cancellationToken) {
        Encoder encoder = YamlEncodingDetector.CreateEncoding(_encoding).GetEncoder();
        char[] chars = ArrayPool<char>.Shared.Rent(EncodeChunk);
        byte[] bytes = ArrayPool<byte>.Shared.Rent(EncodeChunk * 4 + 8);
        try {
            if (_byteOrderMark) {
                byte[] mark = YamlEncodingDetector.GetByteOrderMark(_encoding);
                await stream.WriteAsync(mark, 0, mark.Length, cancellationToken).ConfigureAwait(false);
            }
            for (int s = 0; s < segments.Count; s++) {
                YamlTextSegment segment = segments[s];
                for (int offset = 0; offset < segment.Length; offset += EncodeChunk) {
                    int written = Encode(encoder, segment, offset, chars, bytes, flush: false);
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

    private static int Encode(Encoder encoder, YamlTextSegment segment, int offset, char[] chars, byte[] bytes, bool flush) {
        int count = Math.Min(EncodeChunk, segment.Length - offset);
        segment.Text.CopyTo(segment.Start + offset, chars, 0, count);
        return encoder.GetBytes(chars, 0, count, bytes, 0, flush);
    }

    private static void RequireWritable(Stream stream) {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
    }
}
