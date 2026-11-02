using DocumentFormat.OpenXml.Packaging;
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
        internal ExcelDocumentReader CreateCellEnumerationReader(ExcelSheet sheet, ExcelReaderOptions options) {
            if (AcquireUnsavedCellsSnapshot(sheet, options) is { } unsaved) return ExcelDocumentReader.Wrap(unsaved, options);
            MaterializeDeferredDataSetImport();
            SpreadsheetDocument document = _spreadSheetDocument;
            ExcelPackageReadSnapshot? snapshot = AcquireReadSnapshot();
            if (snapshot != null) return ExcelDocumentReader.Wrap(snapshot, options);
            ExcelPackageReadSnapshot? baseline = AcquireEditedPackageBaseline();
            return baseline != null
                ? ExcelDocumentReader.WrapEdited(document, baseline, options, tabular: false)
                : ExcelDocumentReader.Wrap(document, options);
        }

        /// <summary>
        /// Creates the reader behind UsedRangeA1, which internal edits also call before they report a
        /// change. They have loaded the worksheet by then, and reading a loaded worksheet costs no load.
        /// </summary>
        internal ExcelDocumentReader CreateUsedRangeReader(ExcelSheet sheet) =>
            sheet.DeferredMetadataWorksheetPart.IsRootElementLoaded
                ? CreateReader()
                : CreateTabularReader(sheet, new ExcelReaderOptions());

        /// <summary>
        /// Creates the reader behind the data reader, table and typed routes. Their fast paths
        /// resolve malformed cells and oversized chunks differently from the SDK route, so each
        /// worksheet that does not qualify keeps that route.
        /// </summary>
        internal ExcelDocumentReader CreateTabularReader(ExcelSheet? sheet, ExcelReaderOptions options) {
            if (sheet != null && AcquireUnsavedCellsSnapshot(sheet, options) is { } unsaved) {
                return ExcelDocumentReader.WrapTabular(unsaved, options, () => {
                    MaterializeDeferredDataSetImport();
                    return ExcelDocumentReader.Wrap(_spreadSheetDocument, options);
                });
            }

            MaterializeDeferredDataSetImport();
            SpreadsheetDocument document = _spreadSheetDocument;
            if (options.CellValueConverter != null) return ExcelDocumentReader.Wrap(document, options);
            ExcelPackageReadSnapshot? snapshot = AcquireReadSnapshot();
            if (snapshot != null) {
                return ExcelDocumentReader.WrapTabular(snapshot, options, () => ExcelDocumentReader.Wrap(_spreadSheetDocument, options));
            }

            ExcelPackageReadSnapshot? baseline = AcquireEditedPackageBaseline();
            return baseline != null
                ? ExcelDocumentReader.WrapEdited(document, baseline, options, tabular: true)
                : ExcelDocumentReader.Wrap(document, options);
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
                    || HasPendingDirectCellValues) {
                    return null;
                }

                return _unchangedPackageBytes
                    ?? (_usesFileBackedPackage && _ownedOpenStream != null ? _fileBackedPackageVersion : null);
            }
        }

        /// <summary>
        /// The package an edited document last matched. The document only changes a worksheet through its
        /// root, but exposed SDK access can feed a part directly, so then a part's state says nothing about these bytes.
        /// </summary>
        private byte[]? EditedPackageBaseline =>
            _readSnapshotAllowed && !_openXmlDocumentExposed ? _packageBaselineBytes : null;

        private ExcelPackageReadSnapshot? AcquireReadSnapshot() => AcquireSnapshot(ReadSnapshotIdentity);

        private ExcelPackageReadSnapshot? AcquireEditedPackageBaseline() => AcquireSnapshot(EditedPackageBaseline);

        /// <summary>Moves a snapshot off an unchanged document, unless it stays useful as the edited document's baseline.</summary>
        private void ReleaseReadSnapshotOnEdit() {
            if (!ReferenceEquals(_readSnapshot?.Identity, EditedPackageBaseline)) ReleaseReadSnapshot();
        }

        private ExcelPackageReadSnapshot? AcquireSnapshot(object? identity) {
            lock (_readSnapshotGate) {
                if (identity == null) {
                    ReleaseReadSnapshotOnEdit();
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
                ReleaseUnsavedCellsSnapshot();
                _readSnapshot?.Release();
                _readSnapshot = null;
                _readSnapshotDeclinedIdentity = null;
            }
        }
    }
}
