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

        private void MaterializeDirectDataSetFastSaveModelIfNeeded() {
            var model = _materializedDirectDataSetFastSaveModel;
            if (model == null || _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet) {
                return;
            }

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(model, preserveExistingWorksheetStructure: true);
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
                CombineOverlayCells(metadata.OverlayCells, overlayCells));
        }
    }
}
