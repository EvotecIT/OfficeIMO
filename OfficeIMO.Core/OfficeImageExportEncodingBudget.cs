using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Coordinates an encoded-byte ceiling across one or more concurrent image encoders.</summary>
internal sealed class OfficeImageExportEncodingBudget {
    private readonly SemaphoreSlim _serializedEncodingGate = new SemaphoreSlim(1, 1);
    private long _usedBytes;

    internal OfficeImageExportEncodingBudget(long maximumBytes) {
        if (maximumBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        MaximumBytes = maximumBytes;
    }

    internal long MaximumBytes { get; }

    internal byte[] EncodeWithinRemainingBudget(
        Func<long, byte[]> encoder,
        CancellationToken cancellationToken) {
        if (encoder == null) throw new ArgumentNullException(nameof(encoder));
        cancellationToken.ThrowIfCancellationRequested();
        _serializedEncodingGate.Wait(cancellationToken);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            long used = Volatile.Read(ref _usedBytes);
            long remaining = MaximumBytes - used;
            if (remaining < 1L) {
                throw new OfficeImageExportBatchLimitException(
                    nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                    used == long.MaxValue ? long.MaxValue : used + 1L,
                    MaximumBytes);
            }
            byte[] bytes;
            try {
                bytes = encoder(remaining);
            } catch (OfficeImageExportBatchLimitException exception) {
                long actual = used > long.MaxValue - exception.Actual
                    ? long.MaxValue
                    : used + exception.Actual;
                throw new OfficeImageExportBatchLimitException(
                    nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                    actual,
                    MaximumBytes);
            }
            cancellationToken.ThrowIfCancellationRequested();
            Reserve(bytes.Length);
            return bytes;
        } finally {
            _serializedEncodingGate.Release();
        }
    }

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
