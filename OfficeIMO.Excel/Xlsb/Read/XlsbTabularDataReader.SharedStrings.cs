namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private readonly XlsbSharedStringTable? _sharedStringOwner;
        private bool _hasSharedStringReference;

        private void AcquireSharedStringReference() {
            if (_sharedStringOwner == null) return;
            _sharedStringOwner.AddReaderReference();
            _hasSharedStringReference = true;
        }

        private void ReleaseSharedStringReference() {
            if (!_hasSharedStringReference) return;
            _hasSharedStringReference = false;
            _sharedStringOwner!.ReleaseReaderReference();
        }

        private void StoreSharedStringText(int ordinal, int index) {
#if NET8_0_OR_GREATER
            // Only readers created with borrowing enabled may defer canonical projection.
            _strings[ordinal] = _borrowableText.Length != 0 && !_options.InferSchema
                && _options.CellValueConverter == null ? null : _sharedStrings[index];
#else
            _strings[ordinal] = _sharedStrings[index];
#endif
            TrackUtf8SharedStringCell(ordinal, index);
        }
    }
}
