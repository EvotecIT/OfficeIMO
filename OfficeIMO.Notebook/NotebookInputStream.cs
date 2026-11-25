namespace OfficeIMO.Notebook;

// Lets the JSON parser own the single retained stream buffer. The wrapper never owns the caller's stream.
internal sealed class NotebookInputStream : Stream {
    private readonly Stream _source;
    private readonly long _limit;
    private readonly CancellationToken _cancellationToken;
    private long _remaining;

    internal NotebookInputStream(Stream source, long limit, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(source);
        if (!source.CanRead) throw new ArgumentException("The notebook source must be readable.", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (source.CanSeek && source.Length - source.Position > limit) throw InputLimit(limit);
        _source = source;
        _limit = limit;
        _remaining = limit;
        _cancellationToken = cancellationToken;
    }

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) {
        _cancellationToken.ThrowIfCancellationRequested();
        return RecordRead(_source.Read(buffer, offset, ReadCount(count)));
    }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        _cancellationToken.ThrowIfCancellationRequested();
        return RecordRead(await _source.ReadAsync(buffer, offset, ReadCount(count), cancellationToken).ConfigureAwait(false));
    }

#if NET8_0_OR_GREATER
    public override int Read(Span<byte> buffer) {
        _cancellationToken.ThrowIfCancellationRequested();
        return RecordRead(_source.Read(buffer.Slice(0, ReadCount(buffer.Length))));
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) {
        _cancellationToken.ThrowIfCancellationRequested();
        return RecordRead(await _source.ReadAsync(buffer.Slice(0, ReadCount(buffer.Length)), cancellationToken).ConfigureAwait(false));
    }
#endif

    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    internal static InvalidDataException InputLimit(long limit) =>
        new($"Notebook input exceeds the configured MaxInputBytes limit of {limit.ToString(CultureInfo.InvariantCulture)} bytes.");

    private int ReadCount(int requested) => _remaining < requested ? (int)_remaining + 1 : requested;

    private int RecordRead(int read) {
        _remaining -= read;
        if (_remaining < 0) throw InputLimit(_limit);
        return read;
    }
}
