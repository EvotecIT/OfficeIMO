using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Applies cancellation and an aggregate byte budget to encoder writes.</summary>
internal sealed class OfficeImageExportEncodingStream : Stream {
    private readonly Stream _destination;
    private readonly OfficeImageExportEncodingBudget _budget;
    private readonly CancellationToken _cancellationToken;

    internal OfficeImageExportEncodingStream(
        Stream destination,
        OfficeImageExportEncodingBudget budget,
        CancellationToken cancellationToken) {
        _destination = destination ?? throw new ArgumentNullException(nameof(destination));
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
        _cancellationToken = cancellationToken;
        OfficeRasterOutput.EnsureWritable(destination);
    }

    internal Stream WrappedDestination => _destination;

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() {
        _cancellationToken.ThrowIfCancellationRequested();
        _destination.Flush();
    }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        _cancellationToken.ThrowIfCancellationRequested();
        _budget.Reserve(count);
        _destination.Write(buffer, offset, count);
    }

#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) {
        _cancellationToken.ThrowIfCancellationRequested();
        _budget.Reserve(buffer.Length);
        _destination.Write(buffer);
    }
#endif

    public override void WriteByte(byte value) {
        _cancellationToken.ThrowIfCancellationRequested();
        _budget.Reserve(1);
        _destination.WriteByte(value);
    }
}
