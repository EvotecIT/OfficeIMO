using System;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>Bounds metadata rewrite output, retained inputs, backing growth, and final array materialization.</summary>
internal sealed class OfficeMetadataRewriteStream : OfficeBoundedMemoryStream {
    private long _retainedBytes;
    private readonly CancellationToken _token;

    internal OfficeMetadataRewriteStream(long retainedBytes, int capacityHint, CancellationToken token,
        int maximumBytes = OfficeRasterGuards.MaximumEncodedBytes)
        : base(maximumBytes, ValidateInitialCapacity(retainedBytes, capacityHint, maximumBytes)) {
        _retainedBytes = retainedBytes;
        _token = token;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        _token.ThrowIfCancellationRequested();
        EnsurePeak(count);
        base.Write(buffer, offset, count);
    }
    public override void WriteByte(byte value) { _token.ThrowIfCancellationRequested(); EnsurePeak(1); base.WriteByte(value); }
#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) { _token.ThrowIfCancellationRequested(); EnsurePeak(buffer.Length); base.Write(buffer); }
#endif
    public override byte[] ToArray() {
        _token.ThrowIfCancellationRequested();
        EnsurePeak(0, materialize: true);
        return base.ToArray();
    }
    internal void CheckTransientBytes(long bytes) {
        if (bytes < 0L || checked(_retainedBytes + OfficeRasterOutput.GetMemoryStreamBackingBytes(this) + bytes + 24L) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
    }
    internal void AddRetainedBytes(long bytes) { CheckTransientBytes(bytes); _retainedBytes = checked(_retainedBytes + bytes); }
    private void EnsurePeak(long appendedBytes, bool materialize = false) {
        if (checked(_retainedBytes + OfficeRasterOutput.GetMemoryStreamSingleWritePeakBytes(this, appendedBytes, materialize)) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
    }
    private static int ValidateInitialCapacity(long retainedBytes, int hint, int maximumBytes) {
        if (retainedBytes < 0 || hint < 0) throw new ArgumentOutOfRangeException(nameof(retainedBytes));
        int capacity = Math.Min(hint, maximumBytes);
        if (checked(retainedBytes + capacity + 24L) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
        return capacity;
    }
}
