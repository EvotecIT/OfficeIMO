using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        // MS-XLS 2.5.198.118 prohibits cross-sheet, range-operator and several
        // opaque operand forms in ShrFmla. Emit those OOXML shared cells as
        // ordinary independently relocated Formula records instead.
        private static bool CanWriteSharedParsedFormula(byte[] tokens) {
            var scanned = new List<LegacyXlsFormulaTokenRecord>();
            BiffFormulaTokenScanner.ScanTokens(tokens, "SharedFormula", null, null, 0, 0x04bc, scanned);
            return scanned.Count > 0 && scanned.All(item => item.OperandByteCount != null
                && item.Token is not (0x01 or 0x02 or 0x0f or 0x10 or 0x11 or 0x18)
                && (item.Token < 0x20 || (item.Token & 0x1f) is not (0x00 or 0x06 or 0x07 or 0x08 or 0x09 or 0x0a or 0x0b or 0x19 or 0x1a or 0x1b or 0x1c or 0x1d)))
                && scanned.Last().TokenOffset + 1 + scanned.Last().OperandByteCount == tokens.Length;
        }

        private sealed class LegacyXlsSharedFormulaTable {
            private readonly Dictionary<uint, LegacyXlsSharedFormulaDefinition> _definitions;

            private LegacyXlsSharedFormulaTable(Dictionary<uint, LegacyXlsSharedFormulaDefinition> definitions) {
                _definitions = definitions;
            }

            internal static LegacyXlsSharedFormulaTable Create(ExcelSheet sheet, int sheetIndex, LegacyXlsFormulaNameIndex formulaNameIndex) {
                SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
                var definitions = new Dictionary<uint, LegacyXlsSharedFormulaDefinition>();
                if (sheetData == null) {
                    return new LegacyXlsSharedFormulaTable(definitions);
                }

                foreach (Row row in sheetData.Elements<Row>()) {
                    uint rowIndex = row.RowIndex?.Value ?? 0U;
                    int sequentialColumn = 1;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        CellFormula? formula = cell.CellFormula;
                        if (formula?.FormulaType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared) {
                            if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                                ParseCellReference(cell.CellReference!.Value!, out _, out sequentialColumn);
                            }

                            sequentialColumn++;
                            continue;
                        }

                        uint effectiveRow = rowIndex;
                        int effectiveColumn = sequentialColumn;
                        if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                            ParseCellReference(cell.CellReference!.Value!, out effectiveRow, out effectiveColumn);
                        }

                        sequentialColumn = effectiveColumn + 1;
                        if (effectiveRow == 0 || effectiveColumn <= 0 || formula.SharedIndex == null || string.IsNullOrWhiteSpace(formula.Text)) {
                            continue;
                        }

                        uint sharedIndex = formula.SharedIndex.Value;
                        if (definitions.ContainsKey(sharedIndex)) {
                            throw new NotSupportedException($"Native XLS saving found duplicate shared formula definition index {sharedIndex} on worksheet '{sheet.Name}'. Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        ushort anchorRow = checked((ushort)(effectiveRow - 1U));
                        ushort anchorColumn = checked((ushort)(effectiveColumn - 1));
                        if (!TryParseFormulaRange(formula.Reference?.Value, out ushort firstRow, out ushort firstColumn, out ushort lastRow, out ushort lastColumn)
                            || anchorRow < firstRow
                            || anchorRow > lastRow
                            || anchorColumn < firstColumn
                            || anchorColumn > lastColumn) {
                            throw new NotSupportedException($"Native XLS saving does not support shared formula range '{formula.Reference?.Value}' at {ToA1Address(anchorRow, anchorColumn)}. Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        if (!LegacyXlsFormulaEncoder.TryEncode(formula.Text, formulaNameIndex, sheetIndex, out byte[] formulaTokens, out string? reason)) {
                            throw new NotSupportedException($"Native XLS saving does not yet support shared formula '{formula.Text}' at {ToA1Address(anchorRow, anchorColumn)}: {reason} Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        byte[] sharedFormulaTokens = Array.Empty<byte>();
                        bool shared = CanWriteSharedParsedFormula(formulaTokens)
                            && TryBuildSharedFormulaTokens(formulaTokens, anchorRow, anchorColumn, out sharedFormulaTokens);
                        byte[] payload = Array.Empty<byte>();
                        if (shared) {
                            EnsureSupportedSharedFormulaPayloadLength(sharedFormulaTokens, ToA1Address(anchorRow, anchorColumn));
                            EnsureSupportedSharedFormulaCellCount(firstRow, firstColumn, lastRow, lastColumn, ToA1Address(anchorRow, anchorColumn));
                            payload = BuildSharedFormulaPayload(firstRow, firstColumn, lastRow, lastColumn, sharedFormulaTokens);
                        }
                        definitions[sharedIndex] = new LegacyXlsSharedFormulaDefinition(
                            anchorRow,
                            anchorColumn,
                            formula.Text,
                            payload);
                    }
                }

                return new LegacyXlsSharedFormulaTable(definitions);
            }

            internal bool TryGetDefinition(uint? sharedIndex, out LegacyXlsSharedFormulaDefinition definition) {
                definition = default;
                return sharedIndex.HasValue && _definitions.TryGetValue(sharedIndex.Value, out definition);
            }

            private static void EnsureSupportedSharedFormulaCellCount(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, string address) {
                if (GetSharedFormulaCellCount(firstRow, firstColumn, lastRow, lastColumn) > byte.MaxValue) {
                    throw new NotSupportedException($"Native XLS saving does not yet support shared formula ranges with more than 255 cells at {address}. Save as .xlsx or split this shared formula before saving as .xls.");
                }
            }
        }

        private readonly struct LegacyXlsSharedFormulaDefinition {
            internal LegacyXlsSharedFormulaDefinition(ushort anchorRow, ushort anchorColumn, string formulaText, byte[] payload) {
                AnchorRow = anchorRow;
                AnchorColumn = anchorColumn;
                FormulaText = formulaText;
                Payload = payload;
            }

            internal ushort AnchorRow { get; }

            internal ushort AnchorColumn { get; }

            internal string FormulaText { get; }

            internal byte[] Payload { get; }
        }

    }
}
