using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly object _readSnapshotGate = new();
        private ExcelPackageReadSnapshot? _readSnapshot;
        private object? _readSnapshotDeclinedIdentity;
        private object _fileBackedPackageVersion = new();
        private bool _readSnapshotAllowed;

        /// <summary>
        /// Creates the reader behind cell enumeration. A converter also sees cells without content,
        /// which only the SDK route visits.
        /// </summary>
        internal ExcelDocumentReader CreateCellEnumerationReader(ExcelReaderOptions options) {
            MaterializeDeferredDataSetImport();
            ExcelPackageReadSnapshot? snapshot = AcquireReadSnapshot();
            return snapshot != null
                ? ExcelDocumentReader.Wrap(snapshot, options)
                : ExcelDocumentReader.Wrap(_spreadSheetDocument, options);
        }

        /// <summary>
        /// Creates the reader behind UsedRangeA1, which internal edits also call before they report a
        /// change. They have loaded the worksheet by then, and reading a loaded worksheet costs no load.
        /// </summary>
        internal ExcelDocumentReader CreateUsedRangeReader(ExcelSheet sheet) =>
            sheet.DeferredMetadataWorksheetPart.IsRootElementLoaded
                ? CreateReader()
                : CreateTabularReader(new ExcelReaderOptions());

        /// <summary>
        /// Creates the reader behind the data reader, table and typed routes. Their fast paths
        /// resolve malformed cells and oversized chunks differently from the SDK route, so each
        /// worksheet that does not qualify keeps that route.
        /// </summary>
        internal ExcelDocumentReader CreateTabularReader(ExcelReaderOptions options) {
            MaterializeDeferredDataSetImport();
            ExcelPackageReadSnapshot? snapshot = options.CellValueConverter == null ? AcquireReadSnapshot() : null;
            return snapshot != null
                ? ExcelDocumentReader.WrapTabular(snapshot, options, () => ExcelDocumentReader.Wrap(_spreadSheetDocument, options))
                : ExcelDocumentReader.Wrap(_spreadSheetDocument, options);
        }

        /// <summary>
        /// The same state lets Save write <see cref="_unchangedPackageBytes"/> verbatim, so those bytes
        /// are the workbook; a file-backed package is its temporary file. Exposed SDK access can change
        /// either without reporting a change.
        /// </summary>
        private object? ReadSnapshotIdentity {
            get {
                if (!_readSnapshotAllowed
                    || _openXmlDocumentExposed
                    || _packageDirty
                    || _sharedStringTableDirty
                    || _directDataSetSaveCandidate != null
                    || _pendingDirectCellValueSheet != null) {
                    return null;
                }

                return _unchangedPackageBytes
                    ?? (_usesFileBackedPackage && _ownedOpenStream != null ? _fileBackedPackageVersion : null);
            }
        }

        private ExcelPackageReadSnapshot? AcquireReadSnapshot() {
            lock (_readSnapshotGate) {
                object? identity = ReadSnapshotIdentity;
                if (identity == null) {
                    ReleaseReadSnapshot();
                    return null;
                }

                if (_readSnapshot != null && !ReferenceEquals(_readSnapshot.Identity, identity)) {
                    ReleaseReadSnapshot();
                }

                if (_readSnapshot == null) {
                    if (ReferenceEquals(_readSnapshotDeclinedIdentity, identity)) return null;
                    _readSnapshot = identity is byte[] bytes
                        ? ExcelPackageReadSnapshot.TryOpen(bytes)
                        : ExcelPackageReadSnapshot.TryOpenCopy(_ownedOpenStream!, identity, CancellationToken.None);
                    if (_readSnapshot == null) {
                        _readSnapshotDeclinedIdentity = identity;
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
                _readSnapshotDeclinedIdentity = null;
            }
        }
    }
}
