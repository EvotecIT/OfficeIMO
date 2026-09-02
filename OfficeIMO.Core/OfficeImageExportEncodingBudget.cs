using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Coordinates an encoded-byte ceiling across one or more concurrent image encoders.</summary>
internal sealed class OfficeImageExportEncodingBudget {
    private long _usedBytes;

    internal OfficeImageExportEncodingBudget(long maximumBytes) {
        if (maximumBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        MaximumBytes = maximumBytes;
    }

    internal long MaximumBytes { get; }

    internal void Reserve(int byteCount) {
        if (byteCount < 0) throw new ArgumentOutOfRangeException(nameof(byteCount));
        if (byteCount == 0) return;

        while (true) {
            long current = Volatile.Read(ref _usedBytes);
            long actual = current > long.MaxValue - byteCount
                ? long.MaxValue
                : current + byteCount;
            if (actual > MaximumBytes) {
                throw new OfficeImageExportBatchLimitException(
                    nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                    actual,
                    MaximumBytes);
            }
            if (Interlocked.CompareExchange(ref _usedBytes, actual, current) == current) return;
        }
    }
}
