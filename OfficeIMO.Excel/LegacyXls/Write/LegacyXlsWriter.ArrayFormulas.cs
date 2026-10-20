using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
        private sealed class LegacyXlsArrayFormulaTable {
            private readonly List<LegacyXlsArrayFormulaDefinition> _definitions;

            private LegacyXlsArrayFormulaTable(List<LegacyXlsArrayFormulaDefinition> definitions) {
                _definitions = definitions;
            }

            internal static LegacyXlsArrayFormulaTable Create(ExcelSheet sheet, int sheetIndex, LegacyXlsFormulaNameIndex formulaNameIndex) {
                SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
                var definitions = new List<LegacyXlsArrayFormulaDefinition>();
                if (sheetData == null) {
                    return new LegacyXlsArrayFormulaTable(definitions);
                }

                Dictionary<ulong, Cell> cellsByAddress = BuildCellAddressMap(sheetData);
                foreach (KeyValuePair<ulong, Cell> entry in cellsByAddress) {
                    Cell cell = entry.Value;
                    CellFormula? formula = cell.CellFormula;
                    if (formula?.FormulaType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Array) {
                        continue;
                    }

                    DecodeCellKey(entry.Key, out ushort anchorRow, out ushort anchorColumn);
                    if (!TryParseFormulaRange(formula.Reference?.Value, out ushort firstRow, out ushort firstColumn, out ushort lastRow, out ushort lastColumn)
                        || anchorRow < firstRow
                        || anchorRow > lastRow
                        || anchorColumn < firstColumn
                        || anchorColumn > lastColumn) {
                        throw new NotSupportedException($"Native XLS saving does not support array formula range '{formula.Reference?.Value}' at {ToA1Address(anchorRow, anchorColumn)}. Save as .xlsx or remove this formula before saving as .xls.");
                    }

                    if (!LegacyXlsFormulaEncoder.TryEncode(formula.Text ?? string.Empty, formulaNameIndex, sheetIndex, out byte[] formulaTokens, out string? reason)) {
                        throw new NotSupportedException($"Native XLS saving does not yet support array formula '{formula.Text}' at {ToA1Address(anchorRow, anchorColumn)}: {reason} Save as .xlsx or remove this formula before saving as .xls.");
                    }

                    if (!ExcelNativeFormulaOperandTypes.TryApply(formulaTokens, binary12: false, ExcelNativeFormulaOperandTypes.Array, out reason))
                        throw new NotSupportedException("The native array formula cannot be typed: " + reason);
                    EnsureSupportedArrayFormulaPayloadLength(formulaTokens, ToA1Address(anchorRow, anchorColumn));
                    var definition = new LegacyXlsArrayFormulaDefinition(
                        anchorRow,
                        anchorColumn,
                        firstRow,
                        firstColumn,
                        lastRow,
                        lastColumn,
                        BuildArrayFormulaPayload(firstRow, firstColumn, lastRow, lastColumn, formulaTokens));
                    EnsureArrayFormulaCells(cellsByAddress, definition);
                    definitions.Add(definition);
                }

                return new LegacyXlsArrayFormulaTable(definitions);
            }

            internal bool TryGetDefinition(ushort row, ushort column, out LegacyXlsArrayFormulaDefinition definition) {
                foreach (LegacyXlsArrayFormulaDefinition candidate in _definitions) {
                    if (candidate.Contains(row, column)) {
                        definition = candidate;
                        return true;
                    }
                }

                definition = default;
                return false;
            }

            private static Dictionary<ulong, Cell> BuildCellAddressMap(SheetData sheetData) {
                var cells = new Dictionary<ulong, Cell>();
                foreach (Row row in sheetData.Elements<Row>()) {
                    uint rowIndex = row.RowIndex?.Value ?? 0U;
                    int sequentialColumn = 1;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        uint effectiveRow = rowIndex;
                        int effectiveColumn = sequentialColumn;
                        if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                            ParseCellReference(cell.CellReference!.Value!, out effectiveRow, out effectiveColumn);
                        }

                        if (effectiveRow > 0 && effectiveColumn > 0 && effectiveRow <= 65536U && effectiveColumn <= 256) {
                            cells[EncodeCellKey(checked((ushort)(effectiveRow - 1U)), checked((ushort)(effectiveColumn - 1)))] = cell;
                        }

                        sequentialColumn = effectiveColumn + 1;
                    }
                }

                return cells;
            }

            private static void EnsureArrayFormulaCells(IReadOnlyDictionary<ulong, Cell> cellsByAddress, LegacyXlsArrayFormulaDefinition definition) {
                for (int row = definition.FirstRow; row <= definition.LastRow; row++) {
                    for (int column = definition.FirstColumn; column <= definition.LastColumn; column++) {
                        ushort currentRow = checked((ushort)row);
                        ushort currentColumn = checked((ushort)column);
                        if (!cellsByAddress.TryGetValue(EncodeCellKey(currentRow, currentColumn), out Cell? cell)) {
                            throw new NotSupportedException($"Native XLS saving requires every member of array formula range {definition.Range}. Missing cell at {ToA1Address(currentRow, currentColumn)}.");
                        }

                        if ((currentRow != definition.AnchorRow || currentColumn != definition.AnchorColumn) && cell.CellFormula != null) {
                            throw new NotSupportedException($"Native XLS saving does not support non-anchor formulas inside array formula range {definition.Range}. Remove the formula at {ToA1Address(currentRow, currentColumn)} or save as .xlsx.");
                        }
                    }
                }
            }

            private static ulong EncodeCellKey(ushort row, ushort column) {
                return ((ulong)row << 16) | column;
            }

            private static void DecodeCellKey(ulong key, out ushort row, out ushort column) {
                row = (ushort)(key >> 16);
                column = (ushort)(key & 0xffff);
            }
        }

        private readonly struct LegacyXlsArrayFormulaDefinition {
            internal LegacyXlsArrayFormulaDefinition(ushort anchorRow, ushort anchorColumn, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] payload) {
                AnchorRow = anchorRow;
                AnchorColumn = anchorColumn;
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
                Payload = payload;
            }

            internal ushort AnchorRow { get; }

            internal ushort AnchorColumn { get; }

            internal ushort FirstRow { get; }

            internal ushort FirstColumn { get; }

            internal ushort LastRow { get; }

            internal ushort LastColumn { get; }

            internal byte[] Payload { get; }

            internal string Range => ToA1Address(FirstRow, FirstColumn) + ":" + ToA1Address(LastRow, LastColumn);

            internal bool Contains(ushort row, ushort column) {
                return row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
            }
        }
}
