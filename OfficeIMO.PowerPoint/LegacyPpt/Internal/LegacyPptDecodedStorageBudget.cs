namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Enforces one import-wide limit for decoded picture, OLE, ActiveX,
    /// linked-object, VBA, and round-trip theme payloads retained by the parser.
    /// </summary>
    internal sealed class LegacyPptDecodedStorageBudget {
        private readonly long _maximumDecodedBytes;
        private long _decodedBytes;
        private bool _wasExceeded;

        internal LegacyPptDecodedStorageBudget(long maximumDecodedBytes) {
            if (maximumDecodedBytes < 0) {
                throw new ArgumentOutOfRangeException(
                    nameof(maximumDecodedBytes));
            }
            _maximumDecodedBytes = maximumDecodedBytes;
        }

        internal long DecodedBytes => _decodedBytes;

        internal int RemainingAllocationBytes => checked((int)Math.Min(
            int.MaxValue, Math.Max(0L,
                _maximumDecodedBytes - _decodedBytes)));

        internal void ThrowIfExceeded() {
            if (_wasExceeded) throw CreateLimitException();
        }

        internal void Consume(int byteCount) {
            if (byteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(byteCount));
            }
            if (_decodedBytes > _maximumDecodedBytes - byteCount) {
                _wasExceeded = true;
                throw CreateLimitException();
            }
            _decodedBytes += byteCount;
        }

        internal void RejectAllocation() {
            _wasExceeded = true;
            throw CreateLimitException();
        }

        private InvalidDataException CreateLimitException() =>
            new InvalidDataException(
                "The aggregate decoded embedded-storage size exceeds "
                + $"{_maximumDecodedBytes} bytes.");
    }
}
