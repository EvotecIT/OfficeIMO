using System;
using System.IO;

namespace OfficeIMO.Provenance;

/// <summary>Bounds in-memory provenance serialization before a write can grow the backing buffer.</summary>
internal sealed class OfficeProvenanceBoundedMemoryStream : MemoryStream {
    private readonly long _maximumBytes;

    internal OfficeProvenanceBoundedMemoryStream(long maximumBytes, int capacityHint = 0)
        : base(GetInitialCapacity(maximumBytes, capacityHint)) {
        _maximumBytes = maximumBytes;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWrite(count);
        base.Write(buffer, offset, count);
    }

    public override void WriteByte(byte value) {
        EnsureWrite(1);
        base.WriteByte(value);
    }

#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) {
        EnsureWrite(buffer.Length);
        base.Write(buffer);
    }
#endif

    public override void SetLength(long value) {
        OfficeProvenanceBinary.EnsureOutputWithinLimit(value, _maximumBytes);
        base.SetLength(value);
    }

    private void EnsureWrite(int count) {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        long end;
        try {
            end = checked(Position + count);
        } catch (OverflowException) {
            throw OfficeProvenanceLimitException.CreateOutput(
                $"The rewritten asset exceeds the configured output limit of {_maximumBytes} bytes.");
        }
        OfficeProvenanceBinary.EnsureOutputWithinLimit(Math.Max(Length, end), _maximumBytes);
    }

    private static int GetInitialCapacity(long maximumBytes, int capacityHint) {
        if (maximumBytes <= 0 || maximumBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        }
        if (capacityHint < 0) throw new ArgumentOutOfRangeException(nameof(capacityHint));
        return (int)Math.Min(maximumBytes, capacityHint);
    }
}
