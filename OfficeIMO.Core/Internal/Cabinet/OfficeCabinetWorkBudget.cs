using System;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Shared, finite processing budget for all CAB frames decoded by one operation.</summary>
internal sealed class OfficeCabinetWorkBudget {
    internal const long DefaultMaximumBytes = 2L * 1024 * 1024 * 1024;
    private long _remainingBytes;

    internal OfficeCabinetWorkBudget(long maximumBytes = DefaultMaximumBytes) {
        if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _remainingBytes = maximumBytes;
    }

    /// <summary>Charges a frame before checksumming, buffering or decoding it. The 32 KiB
    /// minimum bounds frame setup; input bytes and MSZIP dictionary replay can increase the charge.
    /// Units describe bounded processing work, rather than exact CPU time or newly expanded content.</summary>
    internal void ConsumeFrame(int compressedLength, int expandedLength, int reserveLength, int dictionaryLength, long offset) {
        long recordBytes = 8L + reserveLength + compressedLength;
        long decodedBytes = (long)expandedLength + dictionaryLength;
        long workBytes = Math.Max(OfficeMsZipDecoder.WindowBytes, Math.Max(recordBytes, decodedBytes));
        if (workBytes > _remainingBytes) {
            throw new OfficeCabinetException("CAB_WORK_LIMIT", "CAB frame processing exceeds the configured aggregate decompression-work byte limit.", offset);
        }
        _remainingBytes -= workBytes;
    }
}
