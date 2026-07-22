namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>Tracks aggregate decoded OfficeArt image bytes for one XLS import.</summary>
    internal sealed class LegacyXlsDecodedImageBudget {
        private readonly int _maximumBytes;
        private int _decodedBytes;

        internal LegacyXlsDecodedImageBudget(int maximumBytes) {
            if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            _maximumBytes = maximumBytes;
        }

        internal int RemainingBytes => _maximumBytes - _decodedBytes;

        internal void Consume(int bytes) {
            if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
            _decodedBytes = checked(_decodedBytes + bytes);
            if (_decodedBytes > _maximumBytes) {
                throw new InvalidDataException(
                    $"The legacy XLS decoded-image payload exceeds the configured limit of {_maximumBytes} bytes.");
            }
        }
    }
}
