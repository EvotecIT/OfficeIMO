namespace OfficeIMO.DjVu;

internal sealed class DjVuReadBudget {
    internal readonly DjVuReadOptions Options;
    internal readonly CancellationToken Cancellation;
    private long _expanded;
    private int _chunks;
    private int _textCharacters;
    private int _textZones;
    private long _retainedBytes;
    private long _iw44CoefficientSamples;
    private long _jb2DecodedSamples;
    private long _jb2CommentBytes;

    internal DjVuReadBudget(DjVuReadOptions options, CancellationToken cancellation) {
        Options = options;
        Cancellation = cancellation;
    }

    internal void Expanded(int count) {
        if (count < 0 || count > Options.MaxExpandedBytes - _expanded) throw new DjVuResourceLimitException(nameof(Options.MaxExpandedBytes));
        _expanded += count;
    }

    internal void Chunk(int depth) {
        Cancellation.ThrowIfCancellationRequested();
        if (depth > Options.MaxDepth) throw new DjVuResourceLimitException(nameof(Options.MaxDepth));
        if (++_chunks > Options.MaxChunks) throw new DjVuResourceLimitException(nameof(Options.MaxChunks));
    }

    internal void TextCharacters(int count) {
        if (count < 0 || count > Options.MaxTextCharacters - _textCharacters) throw new DjVuResourceLimitException(nameof(Options.MaxTextCharacters));
        _textCharacters += count;
    }

    internal void TextZone(int depth) {
        Cancellation.ThrowIfCancellationRequested();
        if (depth > Options.MaxTextZoneDepth) throw new DjVuResourceLimitException(nameof(Options.MaxTextZoneDepth));
        if (++_textZones > Options.MaxTextZones) throw new DjVuResourceLimitException(nameof(Options.MaxTextZones));
    }

    internal void WorkingBytes(long bytes) {
        if (bytes < 0 || bytes > Options.MaxCodecBytes - _retainedBytes) throw new DjVuResourceLimitException(nameof(Options.MaxCodecBytes));
    }

    internal void RetainBytes(long bytes) { WorkingBytes(bytes); _retainedBytes += bytes; }
    internal void ReleaseBytes(long bytes) {
        if (bytes < 0 || bytes > _retainedBytes) throw new InvalidOperationException("Invalid DjVu working-buffer release.");
        _retainedBytes -= bytes;
    }
    internal long RetainedBytes => _retainedBytes;

    internal void Iw44CoefficientSamples(long count) {
        if (count < 0 || count > Options.MaxIw44CoefficientSamples - _iw44CoefficientSamples)
            throw new DjVuResourceLimitException(nameof(Options.MaxIw44CoefficientSamples));
        _iw44CoefficientSamples += count;
    }

    internal void Jb2DecodedSamples(long count) {
        if (count < 0 || count > Options.MaxJb2DecodedSamples - _jb2DecodedSamples)
            throw new DjVuResourceLimitException(nameof(Options.MaxJb2DecodedSamples));
        _jb2DecodedSamples += count;
    }

    internal void Jb2CommentBytes(long count) {
        if (count < 0 || count > Options.MaxJb2CommentBytes - _jb2CommentBytes)
            throw new DjVuResourceLimitException(nameof(Options.MaxJb2CommentBytes));
        _jb2CommentBytes += count;
    }
}
