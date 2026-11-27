using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>The cells of the sheet that hold a formula element, in document order.</summary>
        internal IEnumerable<ExcelFormulaCell> EnumerateFormulaCells() {
            foreach (Cell cell in ExcelWorksheetScan.FormulaCells(WorksheetRoot)) yield return cell;
        }

        private IReadOnlyList<string> ResolveSharedFormulaTextsForStructuralValidation() {
            List<Cell> sharedCells = FindSharedFormulaCells();
            if (sharedCells.Count == 0) {
                return Array.Empty<string>();
            }

            IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions =
                BuildSharedFormulaDefinitions(effectiveCoordinates);

            var resolved = new List<string>();
            foreach (Cell cell in sharedCells) {
                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {ExcelWorksheetScan.FindFormula(cell)!.SharedIndex!.Value} cannot be validated safely.");
                }
                resolved.Add(text);
            }
            return resolved;
        }

        private List<Cell> FindSharedFormulaCells() {
            var cells = new List<Cell>();
            foreach (CellFormula formula in ExcelWorksheetScan.CellFormulas(WorksheetRoot)) {
                if (formula.FormulaType?.Value == CellFormulaValues.Shared
                    && formula.SharedIndex?.Value != null
                    && formula.Parent is Cell cell) {
                    cells.Add(cell);
                }
            }
            return cells;
        }

        private void MaterializeWorkbookSharedFormulasForStructuralEdit() {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart) {
                    continue;
                }

                ExcelSheet sheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(
                        _excelDocument,
                        _spreadSheetDocument,
                        sheetElement,
                        registerSheetWrapper: false);
                sheet.MaterializeSharedFormulasForStructuralEdit();
            }
        }

        private void ValidateWorkbookSharedFormulasForStructuralEdit() {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart) {
                    continue;
                }

                ExcelSheet sheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(
                        _excelDocument,
                        _spreadSheetDocument,
                        sheetElement,
                        registerSheetWrapper: false);
                sheet.ResolveSharedFormulaTextsForStructuralValidation();
            }
        }

        private void MaterializeSharedFormulasForStructuralEdit() {
            List<Cell> sharedCells = FindSharedFormulaCells();
            if (sharedCells.Count == 0) {
                return;
            }

            IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions =
                BuildSharedFormulaDefinitions(effectiveCoordinates);
            var resolved = new List<(CellFormula Formula, string Text)>();
            foreach (Cell cell in sharedCells) {
                CellFormula formula = ExcelWorksheetScan.FindFormula(cell)!;
                if (formula.SharedIndex?.Value is not uint sharedIndex) {
                    continue;
                }

                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {sharedIndex} cannot be materialized safely.");
                }

                resolved.Add((formula, text));
            }

            foreach ((CellFormula formula, string text) in resolved) {
                formula.Text = text;
                formula.FormulaType = null;
                formula.SharedIndex = null;
                formula.Reference = null;
                formula.CalculateCell = true;
            }
        }

        /// <summary>The effective coordinates of every cell of the sheet, including cells that omit their reference.</summary>
        internal IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> BuildEffectiveCellCoordinates() =>
            BuildEffectiveCellCoordinates(null, null);

        internal IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> BuildEffectiveCellCoordinates(
            IEnumerable<Row>? rows,
            IEnumerable<Cell>? cells) {
            var coordinates = new Dictionary<ExcelFormulaCell, (int Row, int Column)>();
            uint previousRow = 0U;
            rows ??= WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>();
            ILookup<OpenXmlElement?, Cell>? cellsByRow =
                cells?.ToLookup(cell => cell.Parent);
            foreach (Row row in rows) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                int previousColumn = 0;
                IEnumerable<Cell> rowCells = cellsByRow == null
                    ? row.Elements<Cell>()
                    : cellsByRow[row];
                foreach (Cell cell in rowCells) {
                    int effectiveColumn = previousColumn + 1;
                    if (ExcelA1.TryParseCellReferenceFast(
                        cell.CellReference?.Value,
                        out int explicitRow,
                        out int explicitColumn)) {
                        coordinates[cell] = (explicitRow, explicitColumn);
                        effectiveColumn = explicitColumn;
                    } else if (effectiveRow <= (uint)ExcelA1.MaxRows
                        && effectiveColumn <= ExcelA1.MaxColumns) {
                        coordinates[cell] = (checked((int)effectiveRow), effectiveColumn);
                    }

                    previousColumn = effectiveColumn;
                }

                previousRow = effectiveRow;
            }

            return coordinates;
        }

    }
}
