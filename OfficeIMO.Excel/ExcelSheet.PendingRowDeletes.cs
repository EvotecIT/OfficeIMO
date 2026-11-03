using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Accepts a row deletion without touching the grid when no check can depend on which row it names.
        /// </summary>
        private bool TryDeferRowDeletion(int firstRow, int count) {
            ExcelDocument.PendingRowDeletes? pending = _excelDocument.CurrentPendingRowDeletes;
            if (pending != null) {
                pending.Add(firstRow, count);
                RemapRowDeletionMetadata(firstRow, count);
                return true;
            }

            if (!CanDeferRowDeletions()) return false;
            PreflightRowDeletion(firstRow, count);
            NormalizeImplicitRowIndices(WorksheetRoot.GetFirstChild<SheetData>()!);
            using (_excelDocument.BeginPendingRowDeletes(this, firstRow, count)) {
                RemapRowDeletionMetadata(firstRow, count);
            }
            return true;
        }

        /// <summary>Metadata does not depend on the grid, so it moves at once and only cells and their formulas wait.</summary>
        private void RemapRowDeletionMetadata(int firstRow, int count) {
            int lastRow = firstRow + count - 1;
            RewriteWorkbookCellFormulaReferences(firstRow, -count, lastRow, includeCellFormulas: false);
            RemapDeletedRowMetadata(firstRow, lastRow, -count);
            ResetStructuralMutationCaches();
        }

        /// <summary>
        /// A deletion is row-independent only without tables, pivots, array, data-table or shared formulas,
        /// and nothing else may hold references into the grid: drawings, charts, connections or the exposed package.
        /// </summary>
        private bool CanDeferRowDeletions() {
            WorkbookPart workbookPart = WorkbookPartRoot;
            if (_excelDocument.IsOpenXmlDocumentExposed
                || workbookPart.ChartsheetParts.Any()
                || workbookPart.DialogsheetParts.Any()
                || workbookPart.PivotTableCacheDefinitionParts.Any()
                || workbookPart.ConnectionsPart != null) {
                return false;
            }

            foreach (OpenXmlElement child in WorksheetRoot.ChildElements) {
                if (child is not (SheetProperties or SheetDimension or SheetViews or SheetFormatProperties
                    or Columns or SheetData or PageMargins or DocumentFormat.OpenXml.Spreadsheet.PageSetup or PrintOptions or HeaderFooter or TableParts)) return false;
            }

            WorksheetPart edited = _worksheetPart;
            foreach (WorksheetPart part in workbookPart.WorksheetParts) {
                Worksheet? worksheet = part.Worksheet;
                if (worksheet == null
                    || part.DrawingsPart != null
                    || part.TableDefinitionParts.Any()
                    || part.PivotTableParts.Any()) return false;
                bool isEdited = ReferenceEquals(part, edited);
                foreach (CellFormula formula in ExcelWorksheetScan.CellFormulas(worksheet)) {
                    if (formula.FormulaType != null) return false;
                    if (!isEdited && (formula.Text ?? string.Empty).IndexOf('!') >= 0) return false;
                }
            }
            return true;
        }

        /// <summary>Applies pending deletions in one pass; the result matches applying them one call at a time.</summary>
        internal void ApplyPendingRowDeletes(ExcelDocument.PendingRowDeletes pending) {
            using ExcelWorksheetScan.FormulaScope formulaScope = ExcelWorksheetScan.BeginFormulaScope();
            int[] deleted = pending.DeletedRows.ToArray();
            Worksheet worksheet = WorksheetRoot;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
            OpenXmlElement? next;
            int cursor = 0;
            for (OpenXmlElement? item = sheetData.FirstChild; item != null; item = next) {
                next = item.NextSibling();
                if (item is not Row row) continue;
                int rowIndex = checked((int)row.RowIndex!.Value);
                while (cursor < deleted.Length && deleted[cursor] < rowIndex) cursor++;
                if (cursor < deleted.Length && deleted[cursor] == rowIndex) {
                    // Stored cells hold no value metadata, so a deleted row drops them without building their DOM form.
                    if (row is ExcelStoredRow stored && stored.IsStoreAvailable) stored.ClearStoredCells();
                    foreach (Cell cell in row.Elements<Cell>().ToList()) RemoveCellWithValueMetadataCleanup(cell);
                    row.Remove();
                } else if (cursor > 0) {
                    MoveRow(row, rowIndex - cursor);
                }
            }

            ExcelFormulaAxisShift shift = ExcelFormulaAxisShift.DeletedRows(deleted);
            foreach (CellFormula formula in ExcelWorksheetScan.CellFormulas(worksheet)) {
                string text = formula.Text ?? string.Empty;
                if (text.Length == 0) continue;
                if (!shift.TryRewrite(text, out string rewritten)) {
                    rewritten = text;
                    foreach ((int firstRow, int count) in pending.Operations) {
                        rewritten = RewriteDeletedFormulaReferences(rewritten, firstRow, firstRow + count - 1, -count, Name, rewriteUnqualifiedReferences: true);
                    }
                }
                if (!ReferenceEquals(rewritten, text)) formula.Text = rewritten;
            }

            InvalidateStructuralFormulaResults();
            ResetStructuralMutationCaches();
        }
    }
}
