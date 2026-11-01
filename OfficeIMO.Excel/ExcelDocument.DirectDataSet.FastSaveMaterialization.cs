using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;

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

        // Moves buffered worksheets to the extended writer, which keeps workbook parts the
        // direct package cannot express while the buffered sheets still skip the DOM.
        private void PrepareCellValueBuffersForExtendedSave(ExcelSaveOptions? options) {
            if (options?.DisableFastPackageWriter == true
                || options?.SafePreflight == true
                || options?.SafeRepairDefinedNames == true
                || _materializedDirectDataSetFastSaveModel != null
                || !WorkbookStylesAreDirectWritable()) {
                return;
            }

            if (_directDataSetSaveCandidate?.IsDeferred == true
                && _directDataSetSaveCandidate.Model.Sheets[0].Table.CellValueBuffer != null) {
                PreserveDeferredDataSetFastSaveModelAndClearCandidate();
                return;
            }

            if (!HasPendingDirectCellValues || _directDataSetSaveCandidate != null) {
                return;
            }

            long cellCount = 0;
            var buffers = new List<(ExcelSheet Sheet, ExcelCellValueBuffer Buffer)>();
            foreach (ExcelSheet sheet in _pendingDirectCellValueSheets) {
                cellCount += sheet.PendingCellValueCount;
                if (sheet.TryTakeCellValueBufferForDirectWorksheet() is { } buffer) {
                    buffers.Add((sheet, buffer));
                }
            }

            if (buffers.Count == 0 || cellCount < ExcelSheet.PendingDirectCellValueMinimumCellCount) {
                return;
            }

            var sourceModel = DirectDataSetWorkbookModel.CreateFromCellValueBuffers(buffers, _dateTimeOffsetWriteStrategy, DateSystem);
            if (!TryCreateDirectPackageModel(sourceModel, out DirectDataSetWorkbookModel packageModel, out _, allowDrawings: true)) {
                return;
            }

            for (int i = 0; i < buffers.Count; i++) {
                buffers[i].Sheet.ReleaseCellValueBuffer();
            }

            MaterializePendingDirectCellValueSheetIfNeeded();
            _materializedDirectDataSetFastSaveModel = packageModel;
            _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
            _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
        }

        // The extended writer replaces the stylesheet with the direct one, so a workbook
        // that defined fonts, fills, borders or differential formats keeps the DOM route.
        private bool WorkbookStylesAreDirectWritable() {
            Stylesheet? stylesheet = WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            return stylesheet == null
                || ((stylesheet.Fonts?.ChildElements.Count ?? 0) <= 1
                    && (stylesheet.Fills?.ChildElements.Count ?? 0) <= 2
                    && (stylesheet.Borders?.ChildElements.Count ?? 0) <= 1
                    && (stylesheet.DifferentialFormats?.ChildElements.Count ?? 0) == 0);
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
