namespace OfficeIMO.Word {
    /// <summary>Tracks aggregate certificate decoding and read work for one signature-inspection or validation pass.</summary>
    internal sealed class OfficePackageCertificateByteBudget {
        private readonly long _maximumBytes;
        private long _consumedBytes;

        internal OfficePackageCertificateByteBudget(long maximumBytes) {
            if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            _maximumBytes = maximumBytes;
        }

        internal void Reserve(long byteCount) {
            if (byteCount < 0 || byteCount > _maximumBytes - _consumedBytes) {
                throw new InvalidDataException("The signature certificates exceed the " + _maximumBytes + " byte aggregate certificate limit.");
            }
            _consumedBytes += byteCount;
        }
    }
}
