using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>The sheet's identity for formula caches: the same for every wrapper of one worksheet.</summary>
        internal object FormulaSheetKey => BoundWorksheetPart;

        /// <summary>Whether the cell at the position holds a formula of its own.</summary>
        internal bool HasFormulaCell(int row, int column) =>
            TryGetDomCell(row, column) is { } cell && ExcelWorksheetScan.FindFormula(cell) != null;

        /// <summary>The cell at the position as the formula engine reads it, or an empty handle.</summary>
        internal ExcelFormulaCell GetFormulaCell(int row, int column) => new(TryGetDomCell(row, column));

        /// <summary>Reads a value of the saved pivot table whose location overlaps the range.</summary>
        internal ExcelPivotDataLookup ReadSavedPivotDataAt(int top, int left, int bottom, int right, string dataField,
            Dictionary<string, object?> criteria, out ExcelCellData value) {
            value = null!;
            PivotTablePart? selected = null;
            foreach (var part in _worksheetPart.PivotTableParts) {
                if (!ExcelA1.TryParseRange(part.PivotTableDefinition?.Location?.Reference?.Value ?? "", out int r1, out int c1, out int r2, out int c2)
                    || top > r2 || bottom < r1 || left > c2 || right < c1) continue;
                if (selected != null) return ExcelPivotDataLookup.NotFound;
                selected = part;
            }
            if (selected == null) return ExcelPivotDataLookup.NotFound;
            try {
                value = ReadSavedPivotData(selected, dataField, criteria);
            } catch (NotSupportedException) {
                return ExcelPivotDataLookup.Unsupported;
            }
            return ExcelPivotDataLookup.Found;
        }

        /// <summary>Writes an imported formula with an explicit text result using the OOXML formula-string representation.</summary>
        internal void CellFormulaWithTextCache(int row, int column,
            string formula, string cachedText) {
            ArgumentNullException.ThrowIfNull(formula);
            ArgumentNullException.ThrowIfNull(cachedText);
            CoerceValueHelper.ValidateSharedStringLength(cachedText, nameof(cachedText));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            WriteLock(() => {
                MaterializePendingDirectCellValues();
                MaterializeDeferredDataSetImportIfNeeded();
                Cell cell = GetWritableValueCell(row, column);
                ClearCellValueMetadata(cell);
                SetExistingCellPlainStringValue(cell, cachedText);
                CompleteCellValueMutation(row, column);
                CellFormulaCore(row, column, safeFormula);
            });
        }

        /// <summary>Removes a formula's cached value and value metadata; true when there was one to remove.</summary>
        private bool ClearFormulaCachedValue(ExcelFormulaCell handle) {
            Cell cell = (Cell)handle.Element!;
            if (ExcelWorksheetScan.FindValue(cell) == null && cell.ValueMetaIndex == null) return false;
            cell.CellValue = null;
            ClearCellValueMetadataAttribute(cell);
            return true;
        }

        private void MarkFormulaCellRecalculated(string reference, long mutationVersion) =>
            _excelDocument.MarkFormulaCellRecalculated(_worksheetPart, reference, mutationVersion);

        private void MarkFormulaSheetRecalculated(long mutationVersion) =>
            _excelDocument.MarkFormulaSheetRecalculated(_worksheetPart, mutationVersion);

        private void SetFormulaCachedValue(ExcelFormulaCell handle, FormulaArgumentValue result) {
            Cell cell = (Cell)handle.Element!;
            // Each write goes through the element's own value and type, which allocate nothing when unchanged; the
            // typed setters would replace the value element and build a new attribute value on every pass.
            if (cell.ValueMetaIndex != null) cell.ValueMetaIndex = null;
            if (result.IsError && TryWriteRichFormulaError(cell, result.ErrorCode)) return;
            if (result.IsBoolean) {
                SetFormulaCachedText(cell, result.Number == 0 ? "0" : "1", CellValues.Boolean);
                return;
            }
            if (!result.IsError && result.SourceCellKind == ExcelCellDataKind.Text && result.Text != null) {
                SetFormulaCachedText(cell, result.Text, CellValues.String);
                return;
            }
            if (result.Number.HasValue) {
                SetFormulaCachedText(cell, InvariantNumberText.Get(result.Number.Value), CellValues.Number);
                return;
            }

            if (result.IsError) {
                SetFormulaCachedText(cell, result.ErrorCode ?? "#VALUE!", CellValues.Error);
                return;
            }

            if (result.Text != null) {
                SetFormulaCachedText(cell, result.Text, CellValues.String);
            }
        }

        private static void SetFormulaCachedText(Cell cell, string text, CellValues type) {
            ExcelWorksheetScan.SetValueText(cell, type == CellValues.String
                ? Utilities.ExcelSanitizer.EncodeXString(text) : text);
            if (cell.DataType?.Value != type) cell.DataType = type;
        }

        private bool FormulaTargetsTableHeader(ExcelFormulaCell cell) => FormulaTargetsTableHeader((Cell)cell.Element!);

        private string? ResolveRichValueError(ExcelFormulaCell cell, string? fallback) => ResolveRichValueError((Cell)cell.Element!, fallback);

        // The dominant operands: numeric and shared-text cells need no typed snapshot, boxed value or error-literal check.
        private bool TryResolvePlainNumberCell(int row, int column, out FormulaArgumentValue value) {
            value = default;
            Cell? cell = ProbeCell(row, column, out ExcelStoredRow? storedRow, out int storedIndex);
            if (storedRow != null) return TryResolveStoredPlainNumberCell(storedRow.CellAt(storedIndex), out value);
            if (cell == null) return false;
            string? text = ExcelWorksheetScan.FindValue(cell)?.Text;
            if (cell.DataType == null || cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return false;
                value = new FormulaArgumentValue(number, text, sourceCellKind: ExcelCellDataKind.Number);
                return true;
            }
            if (cell.DataType.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || string.IsNullOrEmpty(text) || !TryParseCellTextSharedStringIndex(text!, out int index)) return false;
            value = new FormulaArgumentValue(null, BuildCellTextSharedStringSnapshot().Get(index) ?? string.Empty,
                sourceCellKind: ExcelCellDataKind.Text);
            return true;
        }
    }
}
