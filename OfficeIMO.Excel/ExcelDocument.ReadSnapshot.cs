namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly object _readSnapshotGate = new();
        private ExcelPackageReadSnapshot? _readSnapshot;
        private byte[]? _readSnapshotDeclinedBytes;
        private bool _readSnapshotAllowed;

        /// <summary>
        /// Creates the reader behind cell enumeration. A converter also sees cells without content,
        /// which only the SDK route visits.
        /// </summary>
        internal ExcelDocumentReader CreateCellEnumerationReader(ExcelReaderOptions options) {
            MaterializeDeferredDataSetImport();
            ExcelPackageReadSnapshot? snapshot = options.CellValueConverter == null ? AcquireReadSnapshot() : null;
            return snapshot != null
                ? ExcelDocumentReader.Wrap(snapshot, options)
                : ExcelDocumentReader.Wrap(_spreadSheetDocument, options);
        }

        /// <summary>
        /// The same state lets Save write <see cref="_unchangedPackageBytes"/> verbatim, so those bytes
        /// are the workbook. Exposed SDK access can change it without reporting a change.
        /// </summary>
        private bool ReadSnapshotMatchesDocument =>
            _readSnapshotAllowed
            && !_openXmlDocumentExposed
            && !_packageDirty
            && !_sharedStringTableDirty
            && _unchangedPackageBytes != null
            && _directDataSetSaveCandidate == null
            && _pendingDirectCellValueSheet == null;

        private ExcelPackageReadSnapshot? AcquireReadSnapshot() {
            lock (_readSnapshotGate) {
                if (!ReadSnapshotMatchesDocument) {
                    ReleaseReadSnapshot();
                    return null;
                }

                byte[] bytes = _unchangedPackageBytes!;
                if (_readSnapshot != null && !ReferenceEquals(_readSnapshot.Bytes, bytes)) {
                    ReleaseReadSnapshot();
                }

                if (_readSnapshot == null) {
                    if (ReferenceEquals(_readSnapshotDeclinedBytes, bytes)) return null;
                    _readSnapshot = ExcelPackageReadSnapshot.TryOpen(bytes);
                    if (_readSnapshot == null) {
                        _readSnapshotDeclinedBytes = bytes;
                        return null;
                    }
                }

                _readSnapshot.AddReference();
                return _readSnapshot;
            }
        }

        /// <summary>A reader still enumerating keeps its own reference and finishes on these bytes.</summary>
        private void ReleaseReadSnapshot() {
            lock (_readSnapshotGate) {
                _readSnapshot?.Release();
                _readSnapshot = null;
                _readSnapshotDeclinedBytes = null;
            }
        }
    }
}
