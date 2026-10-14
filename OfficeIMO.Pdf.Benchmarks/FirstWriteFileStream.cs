using System.Diagnostics;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>
/// Observes the first successful nonempty synchronous write to the actual destination file.
/// This benchmark observer records event timestamps, not total operation timing or durable persistence.
/// </summary>
internal sealed class FirstWriteFileStream : Stream {
    private readonly FileStream _inner;
    private readonly Action _observeFirstWrite;

    internal FirstWriteFileStream(string path, Action observeFirstWrite) {
        _inner = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        _observeFirstWrite = observeFirstWrite;
    }

    internal long FirstWriteCallStartTimestamp { get; private set; }
    internal long FirstWriteCompletionTimestamp { get; private set; }
    internal int FirstWritePayloadBytes { get; private set; }
    internal long TotalWrittenBytes { get; private set; }
    internal long NonemptyWriteCalls { get; private set; }

    public override bool CanRead => false;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => _inner.CanWrite;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }
    public override void Flush() => _inner.Flush();
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void SetLength(long value) => _inner.SetLength(value);
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        long start = count > 0 && FirstWriteCompletionTimestamp == 0L ? Stopwatch.GetTimestamp() : 0L;
        _inner.Write(buffer, offset, count);
        ObserveCompletedWrite(count, start);
    }

    public override void Write(ReadOnlySpan<byte> buffer) {
        long start = buffer.Length > 0 && FirstWriteCompletionTimestamp == 0L ? Stopwatch.GetTimestamp() : 0L;
        _inner.Write(buffer);
        ObserveCompletedWrite(buffer.Length, start);
    }

    public override void WriteByte(byte value) {
        long start = FirstWriteCompletionTimestamp == 0L ? Stopwatch.GetTimestamp() : 0L;
        _inner.WriteByte(value);
        ObserveCompletedWrite(1, start);
    }

    private void ObserveCompletedWrite(int count, long start) {
        if (count == 0) return;
        // Capture completion before counters and the deliberately separate memory observation.
        long completion = start == 0L ? 0L : Stopwatch.GetTimestamp();
        TotalWrittenBytes += count;
        NonemptyWriteCalls++;
        if (start == 0L) return;
        FirstWriteCallStartTimestamp = start;
        FirstWriteCompletionTimestamp = completion;
        FirstWritePayloadBytes = count;
        _observeFirstWrite();
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }
}
