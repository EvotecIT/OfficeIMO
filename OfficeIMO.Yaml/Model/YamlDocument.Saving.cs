using OfficeIMO.Core.Internal;

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
    /// Saves the document to its <see cref="FilePath"/>, in the encoding it was loaded with. Untouched text is copied
    /// from the source; only edited nodes are written anew.
    /// </summary>
    /// <exception cref="InvalidOperationException">The document has no associated path.</exception>
    public void Save() => Save(AssociatedPath());

    /// <summary>
    /// Saves the document to a file in the encoding it was loaded with, writing its byte-order mark when the input had one,
    /// and makes the file its <see cref="FilePath"/>. The file is replaced atomically, so a failed save leaves the old
    /// contents. An unchanged document reproduces the loaded bytes exactly; an edited one changes only the edited nodes' text.
    /// </summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    public void Save(string path, YamlSaveOptions? options = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        List<YamlTextSegment> segments = CollectSegments(options);
        OfficeFileCommit.Write(fullPath, stream => Write(stream, segments));
        _path = fullPath;
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

    /// <summary>Saves the document to its <see cref="FilePath"/> asynchronously; see <see cref="Save()"/>.</summary>
    /// <param name="cancellationToken">Observed between writes.</param>
    public Task SaveAsync(CancellationToken cancellationToken = default) => SaveAsync(AssociatedPath(), null, cancellationToken);

    /// <summary>Saves the document to a file asynchronously; see <see cref="Save(string, YamlSaveOptions?)"/>.</summary>
    /// <param name="path">The file to create or overwrite.</param>
    /// <param name="options">Save settings, or <see langword="null"/> for a minimal-change save.</param>
    /// <param name="cancellationToken">Observed between writes.</param>
    public async Task SaveAsync(string path, YamlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        List<YamlTextSegment> segments = CollectSegments(options);
        await OfficeFileCommit.WriteAsync(fullPath, (stream, token) => WriteAsync(stream, segments, token), cancellationToken: cancellationToken).ConfigureAwait(false);
        _path = fullPath;
    }

    private string AssociatedPath() =>
        _path ?? throw new InvalidOperationException("The document has no associated path; use Save(path) or Save(stream), or create it with Create(path).");

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
#if NET8_0_OR_GREATER
        // One allocation and one copy: a large edited document otherwise allocates its text twice on the large object heap.
        return string.Create(length, segments, static (span, parts) => {
            int offset = 0;
            for (int i = 0; i < parts.Count; i++) {
                YamlTextSegment segment = parts[i];
                segment.Text.AsSpan(segment.Start, segment.Length).CopyTo(span.Slice(offset));
                offset += segment.Length;
            }
        });
#else
        var text = new StringBuilder(length);
        for (int i = 0; i < segments.Count; i++) text.Append(segments[i].Text, segments[i].Start, segments[i].Length);
        return text.ToString();
#endif
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
#if NET8_0_OR_GREATER
        byte[] output = Encode(segments, out int length);
        try {
            stream.Write(output, 0, length);
        } finally {
            ArrayPool<byte>.Shared.Return(output);
        }
#else
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
#endif
    }

    private async Task WriteAsync(Stream stream, List<YamlTextSegment> segments, CancellationToken cancellationToken) {
#if NET8_0_OR_GREATER
        byte[] output = Encode(segments, out int length);
        try {
            await stream.WriteAsync(output.AsMemory(0, length), cancellationToken).ConfigureAwait(false);
        } finally {
            ArrayPool<byte>.Shared.Return(output);
        }
#else
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
#endif
    }

#if NET8_0_OR_GREATER
    /// <summary>
    /// Encodes the whole output, byte-order mark included, into one pooled buffer, so the stream receives one write of
    /// the exact length: a <see cref="MemoryStream"/> then allocates its buffer once instead of doubling it. Segments
    /// end at node boundaries, never inside a surrogate pair, so each encodes on its own.
    /// </summary>
    private byte[] Encode(List<YamlTextSegment> segments, out int length) {
        Encoding encoding = YamlEncodingDetector.CreateEncoding(_encoding);
        byte[]? mark = _byteOrderMark ? YamlEncodingDetector.GetByteOrderMark(_encoding) : null;
        long total = mark?.Length ?? 0;
        for (int i = 0; i < segments.Count; i++) total += encoding.GetByteCount(segments[i].Text.AsSpan(segments[i].Start, segments[i].Length));
        if (total > Array.MaxLength) throw new InvalidOperationException("The document is too large to encode.");
        byte[] output = ArrayPool<byte>.Shared.Rent((int)Math.Max(total, 1));
        try {
            int offset = 0;
            if (mark is not null) {
                mark.CopyTo(output, 0);
                offset = mark.Length;
            }
            for (int i = 0; i < segments.Count; i++) {
                offset += encoding.GetBytes(segments[i].Text.AsSpan(segments[i].Start, segments[i].Length), output.AsSpan(offset));
            }
            length = offset;
            return output;
        } catch {
            ArrayPool<byte>.Shared.Return(output);
            throw;
        }
    }
#endif

#if !NET8_0_OR_GREATER
    private static int Encode(Encoder encoder, YamlTextSegment segment, int offset, char[] chars, byte[] bytes, bool flush) {
        int count = Math.Min(EncodeChunk, segment.Length - offset);
        segment.Text.CopyTo(segment.Start + offset, chars, 0, count);
        return encoder.GetBytes(chars, 0, count, bytes, 0, flush);
    }
#endif

    private static void RequireWritable(Stream stream) {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
    }
}
