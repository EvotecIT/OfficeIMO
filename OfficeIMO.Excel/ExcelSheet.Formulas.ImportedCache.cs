using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Writes an imported formula with an explicit text result using the OOXML formula-string representation.</summary>
        internal void CellFormulaWithTextCache(int row, int column,
            string formula, string cachedText) {
            if (formula is null) throw new ArgumentNullException(nameof(formula));
            if (cachedText is null) throw new ArgumentNullException(nameof(cachedText));
            CoerceValueHelper.ValidateSharedStringLength(cachedText, nameof(cachedText));
            WriteLock(() => {
                MaterializePendingDirectCellValues();
                MaterializeDeferredDataSetImportIfNeeded();
                Cell cell = GetCell(row, column);
                ClearCellValueMetadata(cell);
                SetExistingCellPlainStringValue(cell, cachedText);
                CompleteCellValueMutation(row, column);
                CellFormulaCore(row, column, formula);
            });
        }
    }
}
