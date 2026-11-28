using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The formula model of a worksheet the SDK package holds: its cells are the SDK's cell elements, read and written
    /// in place. It goes once cell values and cell metadata no longer need the SDK model.
    /// </summary>
    public partial class ExcelSheet {
        private SdkFormulaSheetModel? _sdkFormulaModel;

        internal ExcelFormulaSheetModel SdkFormulaModel => _sdkFormulaModel ??= new SdkFormulaSheetModel(this);

        /// <summary>A handle to an SDK cell element of this sheet.</summary>
        internal ExcelFormulaCell WrapFormulaCell(Cell cell) => new(new SdkFormulaCellRecord(this, cell));

        /// <summary>The SDK cell element behind a handle, opening the SDK package for a cell read from the workbook's XML.</summary>
        private Cell SdkCell(ExcelFormulaCell handle) {
            if (handle.Record is SdkFormulaCellRecord sdk) return sdk.Cell;
            if (!handle.Record!.TryGetPosition(out int row, out int column)) throw new InvalidOperationException("The formula cell has no position.");
            _excelDocument.OpenDeferredSdk("Rich formula error on worksheet '" + Name + "'");
            ExcelSheet sheet = _excelDocument.FindSheetForFormulas(Name) ?? this;
            return sheet.GetCell(row, column);
        }

        private bool TryWriteRichFormulaError(ExcelFormulaCell handle, string? error, DynamicArrayPlan? spillPlan = null) =>
            (error == "#CALC!" || error == "#SPILL!") && TryWriteRichFormulaError(SdkCell(handle), error, spillPlan);

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

        /// <summary>The effective coordinates of the given rows' cells, for a structural edit's bounded element list.</summary>
        internal IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> BuildEffectiveCellCoordinates(IEnumerable<Row>? rows, IEnumerable<Cell>? cells) {
            var coordinates = new Dictionary<ExcelFormulaCell, (int Row, int Column)>();
            uint previousRow = 0U;
            rows ??= WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>() ?? Enumerable.Empty<Row>();
            ILookup<OpenXmlElement?, Cell>? cellsByRow = cells?.ToLookup(cell => cell.Parent);
            foreach (Row row in rows) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                int previousColumn = 0;
                foreach (Cell cell in cellsByRow == null ? row.Elements<Cell>() : cellsByRow[row]) {
                    int effectiveColumn = previousColumn + 1;
                    if (ExcelA1.TryParseCellReferenceFast(cell.CellReference?.Value, out int explicitRow, out int explicitColumn)) {
                        coordinates[WrapFormulaCell(cell)] = (explicitRow, explicitColumn);
                        effectiveColumn = explicitColumn;
                    } else if (effectiveRow <= (uint)ExcelA1.MaxRows && effectiveColumn <= ExcelA1.MaxColumns) {
                        coordinates[WrapFormulaCell(cell)] = (checked((int)effectiveRow), effectiveColumn);
                    }
                    previousColumn = effectiveColumn;
                }
                previousRow = effectiveRow;
            }
            return coordinates;
        }

        private sealed class SdkFormulaCellRecord : ExcelFormulaCellRecord {
            private readonly ExcelSheet _sheet;
            internal readonly Cell Cell;

            internal SdkFormulaCellRecord(ExcelSheet sheet, Cell cell) {
                _sheet = sheet;
                Cell = cell;
            }

            private CellFormula? Formula => ExcelWorksheetScan.FindFormula(Cell);

            internal override object Identity => Cell;
            internal override bool HasFormula => Formula != null;
            internal override string? FormulaText => Formula?.Text;

            internal override ExcelFormulaKind FormulaKind {
                get {
                    if (Formula?.FormulaType is not { HasValue: true } formulaType) return ExcelFormulaKind.Normal;
                    CellFormulaValues type = formulaType.Value;
                    if (type == CellFormulaValues.Shared) return ExcelFormulaKind.Shared;
                    if (type == CellFormulaValues.Array) return ExcelFormulaKind.Array;
                    if (type == CellFormulaValues.DataTable) return ExcelFormulaKind.DataTable;
                    return ExcelFormulaKind.Normal;
                }
            }

            internal override uint? SharedIndex => Formula?.SharedIndex is { HasValue: true } value ? value.Value : null;
            internal override string? FormulaReference => Formula?.Reference?.Value;
            internal override bool? CalculateCell => Formula?.CalculateCell is { HasValue: true } flag ? flag.Value : null;
            internal override string? Reference => Cell.CellReference?.Value;
            internal override string? ValueText => ExcelWorksheetScan.FindValue(Cell)?.Text;

            internal override ExcelCellValueType? ValueType {
                get {
                    if (Cell.DataType is not { HasValue: true } dataType) return Cell.DataType == null ? null : ExcelCellValueType.String;
                    CellValues type = dataType.Value;
                    if (type == CellValues.Number) return ExcelCellValueType.Number;
                    if (type == CellValues.Boolean) return ExcelCellValueType.Boolean;
                    if (type == CellValues.Date) return ExcelCellValueType.Date;
                    if (type == CellValues.Error) return ExcelCellValueType.Error;
                    if (type == CellValues.SharedString) return ExcelCellValueType.SharedString;
                    if (type == CellValues.InlineString) return ExcelCellValueType.InlineString;
                    return ExcelCellValueType.String;
                }
            }

            internal override uint? ValueMetaIndex => Cell.ValueMetaIndex is { HasValue: true } value ? value.Value : null;
            internal override uint? CellMetaIndex => Cell.CellMetaIndex is { HasValue: true } value ? value.Value : null;
            internal override bool HasCellMetaIndex => Cell.CellMetaIndex != null;
            internal override uint? StyleIndex => Cell.StyleIndex is { HasValue: true } value ? value.Value : null;
            internal override bool HasInlineString => Cell.InlineString != null;
            internal override bool HasChildren => Cell.HasChildren;

            internal override string InlineText {
                get {
                    InlineString? inline = ExcelWorksheetScan.FindInlineString(Cell);
                    if (inline == null) return string.Empty;
                    if (inline.Text != null) return Utilities.ExcelSanitizer.DecodeXString(inline.Text.Text ?? string.Empty);
                    var builder = new System.Text.StringBuilder();
                    foreach (Run run in inline.Elements<Run>()) {
                        if (run.Text != null) builder.Append(Utilities.ExcelSanitizer.DecodeXString(run.Text.Text));
                    }
                    return builder.ToString();
                }
            }

            internal override bool TryGetPosition(out int row, out int column) => _sheet.TryGetEffectiveCellCoordinates(Cell, out row, out column);

            internal override string? ResolveRichValueError(string? fallback) => _sheet.ResolveRichValueError(Cell, fallback);

            internal override void SetCalculateCell(bool? calculate) {
                if (Formula is not CellFormula formula) return;
                bool? current = formula.CalculateCell is { HasValue: true } flag ? flag.Value : null;
                if (current == calculate && (calculate.HasValue || formula.CalculateCell == null)) return;
                formula.CalculateCell = calculate.HasValue ? new BooleanValue(calculate.Value) : null;
            }

            internal override void SetValue(string? text, ExcelCellValueType? type) {
                // The element's own value and type allocate nothing when unchanged; the typed setters would replace
                // the value element and build a new attribute value on every pass.
                if (text == null) {
                    if (Cell.CellValue != null) Cell.CellValue = null;
                } else {
                    ExcelWorksheetScan.SetValueText(Cell, text);
                }
                CellValues? sdkType = type switch {
                    null => null,
                    ExcelCellValueType.Number => CellValues.Number,
                    ExcelCellValueType.Boolean => CellValues.Boolean,
                    ExcelCellValueType.Date => CellValues.Date,
                    ExcelCellValueType.Error => CellValues.Error,
                    ExcelCellValueType.SharedString => CellValues.SharedString,
                    ExcelCellValueType.InlineString => CellValues.InlineString,
                    _ => CellValues.String,
                };
                if (sdkType == null) {
                    if (Cell.DataType != null) Cell.DataType = null;
                } else if (Cell.DataType?.Value != sdkType.Value) {
                    Cell.DataType = sdkType.Value;
                }
            }

            internal override void ClearInlineString() {
                if (Cell.InlineString != null) Cell.InlineString = null;
            }

            internal override void ClearValueMetaIndex() {
                if (Cell.ValueMetaIndex != null) ClearCellValueMetadataAttribute(Cell);
            }

            internal override void SetFormulaReference(string reference) => Formula!.Reference = reference;

            internal override void SetOwnFormula(string text) {
                CellFormula formula = Formula!;
                formula.Text = text;
                formula.FormulaType = null;
                formula.SharedIndex = null;
                formula.Reference = null;
                formula.CalculateCell = true;
            }
        }

        private sealed class SdkFormulaSheetModel : ExcelFormulaSheetModel {
            private readonly ExcelSheet _sheet;

            internal SdkFormulaSheetModel(ExcelSheet sheet) => _sheet = sheet;

            private ExcelFormulaCellRecord? Wrap(Cell? cell) => cell == null ? null : new SdkFormulaCellRecord(_sheet, cell);

            internal override Uri PartUri => _sheet.BoundWorksheetPart.Uri;

            // A stored cell never holds a formula, so the lookup leaves it where it is.
            internal override ExcelFormulaCellRecord? Find(int row, int column) => Wrap(_sheet.TryGetDomCell(row, column));

            internal override ExcelFormulaCellRecord GetOrCreate(int row, int column) => new SdkFormulaCellRecord(_sheet, _sheet.GetCell(row, column));

            internal override void RemoveIfEmpty(ExcelFormulaCellRecord cell) {
                Cell element = ((SdkFormulaCellRecord)cell).Cell;
                if (element.StyleIndex == null && !element.HasChildren) element.Remove();
            }

            internal override IEnumerable<ExcelFormulaCellRecord> FormulaCells() {
                foreach (Cell cell in ExcelWorksheetScan.FormulaCells(_sheet.WorksheetRoot)) yield return new SdkFormulaCellRecord(_sheet, cell);
            }

            internal override IEnumerable<ExcelFormulaCellRecord> Cells() {
                foreach (Cell cell in ExcelWorksheetScan.Cells(_sheet.WorksheetRoot)) yield return new SdkFormulaCellRecord(_sheet, cell);
            }

            internal override IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> EffectiveCoordinates() =>
                _sheet.BuildEffectiveCellCoordinates(null, null);

            internal override IEnumerable<string> MergedRanges() {
                foreach (MergeCell merge in _sheet.WorksheetRoot.Elements<MergeCells>().SelectMany(m => m.Elements<MergeCell>()))
                    yield return merge.Reference?.Value ?? string.Empty;
            }

            internal override IReadOnlyList<ExcelFormulaTable> Tables() {
                if (!_sheet.HasTableDefinitions) return Array.Empty<ExcelFormulaTable>();
                return _sheet._worksheetPart.TableDefinitionParts.Select(part => part.Table).OfType<Table>().Select(ToFormulaTable).ToList();
            }

            internal override bool HasTableHeaders => _sheet.HasTableDefinitions;

            internal override bool IsTableHeader(int top, int left, int bottom, int right) => _sheet.RangeContainsTableHeader(top, left, bottom, right);

            internal override bool HasPivotTables => _sheet._worksheetPart.PivotTableParts.Any();

            internal override ExcelCellData ReadValue(int row, int column) => _sheet.GetCellValueSnapshot(row, column);

            internal override bool TryReadPlainValue(int row, int column, out double? number, out string? text) {
                number = null;
                text = null;
                Cell? cell = _sheet.ProbeCell(row, column, out ExcelStoredRow? storedRow, out int storedIndex);
                if (storedRow != null) {
                    if (!_sheet.TryResolveStoredPlainNumberCell(storedRow.CellAt(storedIndex), out FormulaArgumentValue stored)) return false;
                    number = stored.Number;
                    text = stored.Text;
                    return true;
                }
                if (cell == null) return false;
                string? value = ExcelWorksheetScan.FindValue(cell)?.Text;
                if (cell.DataType == null || cell.DataType.Value == CellValues.Number) {
                    if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) return false;
                    number = parsed;
                    text = value;
                    return true;
                }
                if (cell.DataType.Value != CellValues.SharedString
                    || string.IsNullOrEmpty(value) || !TryParseCellTextSharedStringIndex(value!, out int index)) return false;
                text = _sheet.BuildCellTextSharedStringSnapshot().Get(index) ?? string.Empty;
                return true;
            }
        }

        internal static void ValidateCellMetadataPart(CellMetadataPart part) => ValidateInCellImageMetadataPart(part, "Cell metadata");

        internal bool SheetRequestsFullCalculationOnLoad =>
            WorksheetRoot.GetFirstChild<SheetCalculationProperties>()?.FullCalculationOnLoad?.Value == true;

        internal static ExcelFormulaTable ToFormulaTableForWorkbook(Table table) => ToFormulaTable(table);

        private static ExcelFormulaTable ToFormulaTable(Table table) => new ExcelFormulaTable(
            table.Name?.Value,
            table.DisplayName?.Value,
            table.Reference?.Value,
            table.HeaderRowCount?.Value ?? 1U,
            table.TotalsRowCount?.Value ?? 1U,
            table.TotalsRowShown?.Value == true,
            (table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()).Select(column => column.Name?.Value).ToList());
    }
}
