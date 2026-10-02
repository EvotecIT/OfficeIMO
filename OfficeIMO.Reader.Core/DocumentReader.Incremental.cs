using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static bool HasIncrementalHandler(string? name, bool pathInput) =>
        GetActiveHandlerRegistry().TryResolve(NormalizeExtension(TryGetExtension(name ?? string.Empty)), out var handler) &&
        (pathInput ? handler.SupportsIncrementalPath : handler.SupportsIncrementalStream);

    internal static IEnumerable<ReaderChunk> EnumerateChunks(string path, ReaderOptions options, CancellationToken token) {
        if (Directory.Exists(path)) {
            foreach (var chunk in ReadDocument(path, options, token).Chunks) yield return chunk;
            yield break;
        }
        ValidateFilePath(path);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, options));
        if (!TryResolvePathHandler(path, options, token, out var handler, out var detection)) {
            throw CreateUnsupportedInputException(path, detection);
        }
        if (!handler.SupportsIncrementalPath) {
            foreach (var chunk in ReadDocument(path, options, token).Chunks) yield return chunk;
            yield break;
        }
        var source = BuildSourceInfoFromPath(path, ShouldComputeSourceHash(handler, options), token);
        foreach (var chunk in handler.ReadPath!(path, options, token)) {
            token.ThrowIfCancellationRequested();
            yield return EnrichChunk(chunk, source, options.ComputeHashes);
        }
    }

    internal static IEnumerable<ReaderChunk> EnumerateChunks(Stream stream, string? name, ReaderOptions options, CancellationToken token) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        string sourceName = NormalizeLogicalSourceName(name, "memory");
        // Forward-only extraction is selected only by an exact registered extension.
        // Content-first routing retains the snapshot-based rich read contract.
        if (options.DetectionMode == ReaderDetectionMode.PreferContent ||
            !GetActiveHandlerRegistry().TryResolve(NormalizeExtension(TryGetExtension(sourceName)), out var handler) ||
            !handler.SupportsIncrementalStream || (!stream.CanSeek && options.ComputeHashes)) {
            foreach (var chunk in ReadDocument(stream, sourceName, options, token).Chunks) yield return chunk;
            yield break;
        }
        long? maximum = ResolveStreamMaxInputBytes(sourceName, options, stream.CanSeek);
        ReaderInputLimitProbe? probe = ResolveStreamInputLimitProbe(sourceName, options);
        long? initialPosition = stream.CanSeek ? stream.Position : null;
        try {
            if (stream.CanSeek) {
                ReaderInputLimits.EnforceSeekableStreamSize(stream, maximum);
                stream.Position = 0;
            }
            using var bounded = new ReaderIncrementalInputStream(stream, maximum, token);
            Stream input = bounded;
            if (probe != null) {
                byte[] prefix = new byte[probe.PrefixLength];
                int length = 0;
                while (length < prefix.Length) {
                    int read = bounded.Read(prefix, length, prefix.Length - length);
                    if (read == 0) break;
                    length += read;
                }
                bounded.ApplyPrefixLimit(probe.ResolveMaxInputBytes(new ReadOnlyMemory<byte>(prefix, 0, length)));
                if (stream.CanSeek) ReaderInputLimits.EnforceSeekableStreamSize(stream, bounded.Maximum);
                input = new ReaderPrefixStream(bounded, prefix, 0, length);
            }
            var source = BuildSourceInfoFromStream(stream, sourceName, ShouldComputeSourceHash(handler, options), token);
            foreach (var chunk in handler.ReadStream!(input, sourceName, options, token)) {
                token.ThrowIfCancellationRequested();
                yield return EnrichChunk(chunk, source, options.ComputeHashes);
            }
        } finally {
            if (initialPosition.HasValue) stream.Position = initialPosition.Value;
        }
    }
}

/// <summary>A forward-only bounded view; disposing it preserves the caller-owned input.</summary>
internal sealed class ReaderIncrementalInputStream : Stream {
    private readonly Stream _source;
    private long? _maximum;
    private readonly CancellationToken _token;
    private long _read;
    private readonly string? _limitName;
    private readonly long? _reportedMaximum;
    internal ReaderIncrementalInputStream(Stream source, long? maximum, CancellationToken token, string? limitName = null, long? reportedMaximum = null) {
        _source = source; _maximum = maximum; _token = token; _limitName = limitName; _reportedMaximum = reportedMaximum;
    }
    internal long? Maximum => _maximum;
    internal void ApplyPrefixLimit(long? maximum) {
        if (maximum.HasValue && maximum.Value < 1) throw new InvalidOperationException("An input-limit prefix resolver returned a value below 1.");
        if (maximum.HasValue) _maximum = _maximum.HasValue ? Math.Min(_maximum.Value, maximum.Value) : maximum;
        if (_maximum.HasValue && _read > _maximum.Value) throw new IOException($"Input exceeds MaxInputBytes ({_maximum.Value}).");
    }
    public override int Read(byte[] buffer, int offset, int count) {
        _token.ThrowIfCancellationRequested();
        long remaining = _maximum.HasValue ? Math.Max(0, _maximum.Value - _read) : long.MaxValue;
        int requested = remaining >= count ? count : (int)remaining + 1;
        int read = _source.Read(buffer, offset, requested);
        _read = checked(_read + read);
        if (_maximum.HasValue && _read > _maximum.Value) {
            if (_limitName != null) throw new ReaderResourceLimitException(_limitName, _reportedMaximum ?? _maximum.Value);
            throw new IOException($"Input exceeds MaxInputBytes ({_maximum.Value}).");
        }
        return read;
    }
    public override bool CanRead => _source.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
