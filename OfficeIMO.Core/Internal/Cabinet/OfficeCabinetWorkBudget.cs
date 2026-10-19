using System;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Shared, finite processing budget for all CAB frames decoded by one operation.</summary>
internal sealed class OfficeCabinetWorkBudget {
    internal const long DefaultMaximumBytes = 2L * 1024 * 1024 * 1024;
    private readonly long _maximumBytes;
    private long _remainingBytes;
    private readonly Action<long, long>? _consumeFrame;
    private readonly Action<long>? _select;

    internal OfficeCabinetWorkBudget(long maximumBytes = DefaultMaximumBytes, Action<long, long>? consumeFrame = null, Action<long>? select = null) {
        _consumeFrame = consumeFrame; _select = select;
        if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
        _remainingBytes = maximumBytes;
    }

    /// <summary>Frame-processing work actually charged, including attempts that fail after charging.</summary>
    internal long ConsumedBytes => _maximumBytes - _remainingBytes;
    internal long MaximumBytes => _maximumBytes;

    internal void Select(long bytes) => _select?.Invoke(bytes);

    /// <summary>Charges a frame before checksumming, buffering or decoding it. The 32 KiB
    /// minimum bounds frame setup; input bytes and MSZIP dictionary replay can increase the charge.
    /// Units describe bounded processing work, rather than exact CPU time or newly expanded content.</summary>
    internal void ConsumeFrame(int compressedLength, int expandedLength, int reserveLength, int dictionaryLength, long offset) {
        long workBytes = FrameWork(compressedLength, expandedLength, reserveLength, dictionaryLength);
        if (workBytes > _remainingBytes) {
            throw new OfficeCabinetException("CAB_WORK_LIMIT", "CAB frame processing exceeds the configured aggregate decompression-work byte limit.", offset);
        }
        _consumeFrame?.Invoke(workBytes, expandedLength);
        _remainingBytes -= workBytes;
    }

    /// <summary>The same deterministic frame charge used by passive planning and actual admission.</summary>
    internal static long FrameWork(int compressedLength, int expandedLength, int reserveLength, int dictionaryLength) {
        long recordBytes = 8L + reserveLength + compressedLength;
        long decodedBytes = (long)expandedLength + dictionaryLength;
        return Math.Max(OfficeMsZipDecoder.WindowBytes, Math.Max(recordBytes, decodedBytes));
    }
}
