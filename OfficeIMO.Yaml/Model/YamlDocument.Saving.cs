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
    /// Saves the document to a file in the encoding it was loaded with, writing its byte-order mark when the input had one.
    /// An unchanged document reproduces the loaded bytes exactly.
    /// </summary>
    /// <param name="path">The file to create or overwrite.</param>
    public void Save(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize);
        Save(stream);
    }

    /// <summary>
    /// Saves the document to a caller-owned stream in the encoding it was loaded with, including its byte-order mark.
    /// The stream is not disposed.
    /// </summary>
    /// <param name="stream">A writable stream.</param>
    public void Save(Stream stream) {
        RequireWritable(stream);
        List<YamlTextSegment> segments = CollectSegments();
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

    /// <summary>Saves the document to a file asynchronously; see <see cref="Save(string)"/>.</summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.Asynchronous);
        await SaveAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the document to a caller-owned stream asynchronously; see <see cref="Save(Stream)"/>.</summary>
    /// <param name="stream">A writable stream. It is not disposed.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public async Task SaveAsync(Stream stream, CancellationToken cancellationToken = default) {
        RequireWritable(stream);
        List<YamlTextSegment> segments = CollectSegments();
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

    /// <summary>
    /// Returns the document as YAML text. An unchanged document returns its source text: the string passed to
    /// <see cref="Parse"/>, or the decoded file or stream without its byte-order mark.
    /// </summary>
    public override string ToString() {
        List<YamlTextSegment> segments = CollectSegments();
        if (segments.Count == 1 && segments[0].Start == 0 && segments[0].Length == segments[0].Text.Length) return segments[0].Text;
        var text = new StringBuilder(_source.Length);
        for (int i = 0; i < segments.Count; i++) text.Append(segments[i].Text, segments[i].Start, segments[i].Length);
        return text.ToString();
    }

    /// <summary>
    /// Lists the output text: source ranges for everything clean, and emitted text for dirty nodes. The walk descends only
    /// into nodes that contain a dirty descendant, so an unchanged document is one range covering the whole source.
    /// </summary>
    internal List<YamlTextSegment> CollectSegments() {
        var segments = new List<YamlTextSegment>(1);
        int cursor = 0;
        for (int i = 0; i < _parts.Count; i++) CollectChanges(_parts[i].Root, segments, ref cursor);
        AddSource(segments, cursor, _source.Length);
        return segments;
    }

    private void CollectChanges(YamlNode node, List<YamlTextSegment> segments, ref int cursor) {
        if ((node._state & YamlNodeState.Dirty) != 0) {
            AddSource(segments, cursor, node._start);
            node.WriteChanged(segments);
            cursor = node._end;
            return;
        }
        if ((node._state & YamlNodeState.ContainsDirty) == 0) return;
        if (node is YamlMapping mapping) {
            for (int i = 0; i < mapping._count; i++) {
                CollectChanges(mapping._entries[i].Key, segments, ref cursor);
                CollectChanges(mapping._entries[i].Value, segments, ref cursor);
            }
        } else if (node is YamlSequence sequence) {
            for (int i = 0; i < sequence._count; i++) CollectChanges(sequence._items[i], segments, ref cursor);
        }
    }

    private void AddSource(List<YamlTextSegment> segments, int start, int end) {
        if (end > start) segments.Add(new YamlTextSegment(_source, start, end - start));
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
