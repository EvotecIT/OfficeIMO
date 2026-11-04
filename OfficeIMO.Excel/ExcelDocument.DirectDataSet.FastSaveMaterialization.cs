using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void PreserveDeferredDataSetFastSaveModelAndClearCandidate() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                ClearDirectDataSetSaveCandidate();
                return;
            }

            if (!TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel packageModel, out _, allowDrawings: true)) {
                return;
            }

            _materializedDirectDataSetFastSaveModel = packageModel;
            _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
            _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
            _directDataSetSaveCandidate = null;
            candidate.Dispose();
        }

        // Moves buffered worksheets to the extended writer, which keeps the workbook stylesheet
        // and every part the direct package cannot express while those sheets skip the DOM.
        private void PrepareCellValueBuffersForExtendedSave(ExcelSaveOptions? options) {
            if (options?.DisableFastPackageWriter == true
                || options?.SafePreflight == true
                || options?.SafeRepairDefinedNames == true
                || _materializedDirectDataSetFastSaveModel != null) {
                return;
            }

            long cellCount = 0;
            var buffers = new List<(ExcelSheet Sheet, ExcelCellValueBuffer Buffer)>();
            DirectDataSetSaveCandidate? candidate = _directDataSetSaveCandidate;
            if (candidate != null) {
                if (!candidate.IsDeferred
                    || candidate.Model.Sheets.Count != 1
                    || candidate.Model.Sheets[0].ColumnNumberFormats != null
                    || candidate.Model.Sheets[0].Table.CellValueBuffer is not { } promoted
                    || TryGetExistingSheet(candidate.Model.Sheets[0].SheetName) is not { } promotedSheet) {
                    return;
                }

                buffers.Add((promotedSheet, promoted));
                cellCount = promoted.CellCount;
            } else {
                foreach (ExcelSheet sheet in _pendingDirectCellValueSheets) {
                    cellCount += sheet.PendingCellValueCount;
                    if (sheet.TryTakeCellValueBufferForDirectWorksheet() is { } buffer) {
                        buffers.Add((sheet, buffer));
                    }
                }
            }

            if (buffers.Count == 0 || cellCount < ExcelSheet.PendingDirectCellValueMinimumCellCount) {
                return;
            }

            var sourceModel = DirectDataSetWorkbookModel.CreateFromCellValueBuffers(
                buffers, _dateTimeOffsetWriteStrategy, DateSystem, buffers[0].Sheet.CreateWorkbookCellStyles());
            if (!TryCreateDirectPackageModel(sourceModel, out DirectDataSetWorkbookModel packageModel, out _, allowDrawings: true)) {
                return;
            }

            for (int i = 0; i < buffers.Count; i++) {
                buffers[i].Sheet.ReleaseCellValueBuffer();
            }

            ClearDirectDataSetSaveCandidate();
            MaterializePendingDirectCellValueSheetIfNeeded();
            _materializedDirectDataSetFastSaveModel = packageModel;
            _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
            _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
        }

        private void MaterializeDirectDataSetFastSaveModelIfNeeded(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var model = _materializedDirectDataSetFastSaveModel;
            if (model == null || _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet) {
                return;
            }

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(
                    model,
                    preserveExistingWorksheetStructure: true,
                    cancellationToken);
                _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = true;
            } finally {
                _materializingDeferredDataSetImport = false;
            }
        }

        private static DirectWorksheetMetadata? MergeDirectMaterializationOverlayCells(
            DirectWorksheetMetadata? metadata,
            IReadOnlyList<DirectOverlayCell> overlayCells) {
            if (overlayCells.Count == 0) {
                return metadata;
            }

            metadata ??= DirectWorksheetMetadata.Empty;
            return new DirectWorksheetMetadata(
                metadata.SheetPropertiesXml,
                metadata.SheetViewsXml,
                metadata.SheetFormatPropertiesXml,
                metadata.SheetProtectionXml,
                metadata.AutoFilterXml,
                metadata.ConditionalFormattingXml,
                metadata.DataValidationsXml,
                metadata.DrawingXml,
                metadata.PostDataValidationXml,
                CombineOverlayCells(metadata.OverlayCells, overlayCells),
                metadata.Structure);
        }
    }
}
