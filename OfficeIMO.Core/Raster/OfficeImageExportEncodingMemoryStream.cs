using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Materializes one encoded image while preserving codec MemoryStream guards, cancellation,
/// aggregate byte accounting, and the final-copy managed-memory ceiling.
/// </summary>
internal sealed class OfficeImageExportEncodingMemoryStream : MemoryStream {
    private readonly OfficeImageExportEncodingBudget _budget;
    private readonly CancellationToken _cancellationToken;
    private readonly long _retainedManagedBytes;

    internal OfficeImageExportEncodingMemoryStream(
        OfficeImageExportEncodingBudget budget,
        CancellationToken cancellationToken,
        long retainedManagedBytes) {
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
        if (retainedManagedBytes < 0L) throw new ArgumentOutOfRangeException(nameof(retainedManagedBytes));
        _cancellationToken = cancellationToken;
        _retainedManagedBytes = retainedManagedBytes;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        _cancellationToken.ThrowIfCancellationRequested();
        EnsureCapacityFor(checked(Position + count));
        _budget.Reserve(count);
        base.Write(buffer, offset, count);
    }

#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) {
        _cancellationToken.ThrowIfCancellationRequested();
        EnsureCapacityFor(checked(Position + buffer.Length));
        _budget.Reserve(buffer.Length);
        base.Write(buffer);
    }
#endif

    public override void WriteByte(byte value) {
        _cancellationToken.ThrowIfCancellationRequested();
        EnsureCapacityFor(checked(Position + 1L));
        _budget.Reserve(1);
        base.WriteByte(value);
    }

    internal byte[] ToBoundedArray() {
        _cancellationToken.ThrowIfCancellationRequested();
        long backingBytes = TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null
            ? segment.Array.LongLength
            : Capacity;
        if (!IsFinalMaterializationWithinLimit(_retainedManagedBytes, backingBytes, Length)) {
            throw new ArgumentException("Image encoding exceeds the managed working-set limit.");
        }
        return ToArray();
    }

    internal static bool IsFinalMaterializationWithinLimit(
        long retainedManagedBytes,
        long backingBytes,
        long encodedLength) {
        if (retainedManagedBytes < 0L || backingBytes < 0L || encodedLength < 0L ||
            encodedLength > OfficeRasterGuards.MaximumEncodedBytes) {
            return false;
        }
        try {
            return checked(
                       retainedManagedBytes + backingBytes + 24L + encodedLength + 24L) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private void EnsureCapacityFor(long requiredPosition) {
        long requiredLength = Math.Max(requiredPosition, Length);
        if (requiredLength > OfficeRasterGuards.MaximumEncodedBytes) {
            throw new ArgumentException("Image encoding exceeds the encoded-size limit.");
        }

        long currentBackingBytes = TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null
            ? segment.Array.LongLength
            : Capacity;
        try {
            long peakBytes;
            if (requiredLength > Capacity) {
                long doubled = Math.Max(256L, checked((long)Capacity * 2L));
                long projectedBackingBytes = Math.Max(requiredLength, doubled);
                peakBytes = checked(
                    _retainedManagedBytes +
                    currentBackingBytes + 24L +
                    projectedBackingBytes + 24L);
            } else {
                peakBytes = checked(_retainedManagedBytes + currentBackingBytes + 24L);
            }
            if (peakBytes > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new ArgumentException("Image encoding exceeds the managed working-set limit.");
            }
        } catch (OverflowException) {
            throw new ArgumentException("Image encoding exceeds the managed working-set limit.");
        }
    }
}
