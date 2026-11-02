using OfficeIMO.Excel.Utilities;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private ExcelPackageReadSnapshot? _unsavedCellsSnapshot;
        private object? _unsavedCellsSource;
        private int _unsavedCellsVersion;

        /// <summary>
        /// Writes a sheet whose cells have not reached the DOM to a private package and reads that, so a read
        /// between writes leaves the cells buffered and sees exactly what a save would store. Returns null when
        /// only the DOM can answer, which the caller reaches by materializing as before.
        /// </summary>
        private ExcelPackageReadSnapshot? AcquireUnsavedCellsSnapshot(ExcelSheet sheet, ExcelReaderOptions options) {
            if (options.CellValueConverter != null || !RequiresDeferredMaterialization || _materializingDeferredDataSetImport) {
                return null;
            }

            return Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)
                ? AcquireUnsavedCellsSnapshotLocked(sheet)
                : Locking.ExecuteWrite(EnsureLock(), () => AcquireUnsavedCellsSnapshotLocked(sheet));
        }

        private ExcelPackageReadSnapshot? AcquireUnsavedCellsSnapshotLocked(ExcelSheet sheet) {
            DirectDataSetSaveCandidate? candidate = _directDataSetSaveCandidate;
            ExcelCellValueBuffer? buffer = candidate == null ? sheet.TryTakeCellValueBufferForDirectWorksheet() : null;
            object? source = candidate != null && IsDeferredModelOf(candidate, sheet) ? candidate : buffer;
            if (source == null) return null;

            int version = buffer?.Version ?? 0;
            lock (_readSnapshotGate) {
                if (_unsavedCellsSnapshot != null && ReferenceEquals(_unsavedCellsSource, source) && _unsavedCellsVersion == version) {
                    _unsavedCellsSnapshot.AddReference();
                    return _unsavedCellsSnapshot;
                }
            }

            if (buffer?.TryReserveReadCopy() == false) return null;

            ExcelPackageReadSnapshot? snapshot = buffer != null
                ? WriteCellValueBufferSnapshot(sheet, buffer)
                : WriteDeferredModelSnapshot(candidate!);
            if (snapshot == null) return null;

            lock (_readSnapshotGate) {
                _unsavedCellsSnapshot?.Release();
                _unsavedCellsSnapshot = snapshot;
                _unsavedCellsSource = source;
                _unsavedCellsVersion = version;
                snapshot.AddReference();
            }

            return snapshot;
        }

        private static bool IsDeferredModelOf(DirectDataSetSaveCandidate candidate, ExcelSheet sheet) {
            if (!candidate.IsDeferred || !candidate.IsValid) return false;
            foreach (DirectDataSetSheetModel model in candidate.Model.Sheets) {
                if (string.Equals(model.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }

        private ExcelPackageReadSnapshot? WriteDeferredModelSnapshot(DirectDataSetSaveCandidate candidate) =>
            TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel model, out _, allowDrawings: true)
                ? WriteUnsavedCellsSnapshot(model, styles: null)
                : null;

        // A format chosen for a cell is an index into the workbook's stylesheet, which the private package does not carry.
        private ExcelPackageReadSnapshot? WriteCellValueBufferSnapshot(ExcelSheet sheet, ExcelCellValueBuffer buffer) {
            DirectWorkbookCellStyles? workbookStyles = buffer.HasCellStyles ? sheet.CreateWorkbookCellStyles() : null;
            var model = DirectDataSetWorkbookModel.CreateFromCellValueBuffers(
                new[] { (sheet, buffer) }, _dateTimeOffsetWriteStrategy, DateSystem, workbookStyles);
            return WriteUnsavedCellsSnapshot(model, workbookStyles != null ? StylesCache.Build(_spreadSheetDocument) : null);
        }

        private ExcelPackageReadSnapshot? WriteUnsavedCellsSnapshot(DirectDataSetWorkbookModel model, StylesCache? styles) {
            using var package = new MemoryStream();
            DirectDataSetWorkbookWriter.Write(package, model, CancellationToken.None, themeXml: null, disableSharedStrings: true);
            return ExcelPackageReadSnapshot.TryOpen(package.ToArray(), styles);
        }

        private void ReleaseUnsavedCellsSnapshot() {
            lock (_readSnapshotGate) {
                _unsavedCellsSnapshot?.Release();
                _unsavedCellsSnapshot = null;
                _unsavedCellsSource = null;
            }
        }
    }
}
