namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Reads newer direct-save widths after deferred state has been materialized by the caller.
        /// The caller holds a workbook read or write lock; the returned model is not mutated.
        /// </summary>
        internal IReadOnlyList<double>? GetDirectColumnWidthsForSnapshot(ExcelSheet sheet) {
            var candidate = _directDataSetSaveCandidate;
            if (!ReferenceEquals(sheet.Document, this)
                || candidate == null || !candidate.IsValid || candidate.IsDeferred) {
                return null;
            }

            foreach (var sheetModel in candidate.Model.Sheets) {
                if (string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                    return sheetModel.ColumnWidths;
                }
            }
            return null;
        }
    }
}
