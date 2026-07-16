using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes worksheet cells and their cached BIFF12 formula values.</summary>
    internal static class XlsbWorksheetCellReader {
        private const int BrtCellBlank = 1;
        private const int BrtCellRk = 2;
        private const int BrtCellError = 3;
        private const int BrtCellBool = 4;
        private const int BrtCellReal = 5;
        private const int BrtCellSt = 6;
        private const int BrtCellIsst = 7;
        private const int BrtFmlaString = 8;
        private const int BrtFmlaNum = 9;
        private const int BrtFmlaBool = 10;
        private const int BrtFmlaError = 11;
        private const int BrtCellRString = 62;

        internal static XlsbCell Read(
            XlsbRecord record,
            int zeroBasedRow,
            IReadOnlyList<string> sharedStrings,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName) {
            var cursor = new XlsbBinaryCursor(record.Data);
            int zeroBasedColumn = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (zeroBasedColumn < 0 || zeroBasedColumn >= A1.MaxColumns) {
                throw new InvalidDataException($"The XLSB cell record at offset {record.Offset} contains invalid column index {zeroBasedColumn}.");
            }

            int row = zeroBasedRow + 1;
            int column = zeroBasedColumn + 1;
            XlsbCell cell;
            switch (record.Type) {
                case BrtCellBlank:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Blank, null, styleIndex);
                    break;
                case BrtCellRk:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Number, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), styleIndex);
                    break;
                case BrtCellError:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Error, BiffErrorValue.ToText(cursor.ReadByte()), styleIndex);
                    break;
                case BrtCellBool:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Boolean, cursor.ReadByte() != 0, styleIndex);
                    break;
                case BrtCellReal:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Number, cursor.ReadDouble(), styleIndex);
                    break;
                case BrtCellSt:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, cursor.ReadWideString(options.MaxStringCharacters), styleIndex);
                    break;
                case BrtCellIsst:
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= sharedStrings.Count) {
                        throw new InvalidDataException($"The XLSB cell at row {row}, column {column} refers to missing shared string {sharedStringIndex}.");
                    }
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, sharedStrings[checked((int)sharedStringIndex)], styleIndex);
                    break;
                case BrtCellRString:
                    byte flags = cursor.ReadByte();
                    string richText = cursor.ReadWideString(options.MaxStringCharacters);
                    if ((flags & 0x03) != 0 || cursor.Remaining > 0) PreserveRecord(options, workbook, partName, record);
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, richText, styleIndex);
                    break;
                case BrtFmlaNum:
                    cell = ReadFormula(record, cursor, row, column, styleIndex, XlsbCellValueKind.Number, cursor.ReadDouble(), options, workbook, partName);
                    break;
                case BrtFmlaBool:
                    cell = ReadFormula(record, cursor, row, column, styleIndex, XlsbCellValueKind.Boolean, cursor.ReadByte() != 0, options, workbook, partName);
                    break;
                case BrtFmlaError:
                    cell = ReadFormula(record, cursor, row, column, styleIndex, XlsbCellValueKind.Error, BiffErrorValue.ToText(cursor.ReadByte()), options, workbook, partName);
                    break;
                case BrtFmlaString:
                    cell = ReadFormula(record, cursor, row, column, styleIndex, XlsbCellValueKind.Text, cursor.ReadWideString(options.MaxStringCharacters), options, workbook, partName);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }

            cell.SourceRecordType = record.Type;
            cell.SourceRecordData = (byte[])record.Data.Clone();
            int availableFormats = workbook.Stylesheet?.CellFormats.Count ?? 1;
            if (styleIndex >= availableFormats) {
                throw new InvalidDataException($"The XLSB cell at row {row}, column {column} refers to missing cell format {styleIndex}; the styles part exposes {availableFormats} format(s).");
            }
            return cell;
        }

        private static XlsbCell ReadFormula(
            XlsbRecord record,
            XlsbBinaryCursor cursor,
            int row,
            int column,
            uint styleIndex,
            XlsbCellValueKind valueKind,
            object? cachedValue,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName) {
            int formulaPayloadOffset = cursor.Position;
            cursor.ReadUInt16(); // grbit flags
            uint tokenCount = cursor.ReadUInt32();
            if (tokenCount > cursor.Remaining) {
                throw new InvalidDataException($"The XLSB formula record at offset {record.Offset} declares {tokenCount} token bytes but only {cursor.Remaining} remain.");
            }

            byte[] tokens = cursor.ReadBytes(checked((int)tokenCount));
            var cell = new XlsbCell(row, column, valueKind, cachedValue, styleIndex) {
                FormulaBytes = tokens,
                FormulaPayloadBytes = CopyTail(record.Data, formulaPayloadOffset)
            };
            if (XlsbFormulaTextReader.TryRead(tokens, out string? formulaText)) {
                cell.FormulaText = formulaText;
            } else {
                workbook.AddDiagnostic(new XlsbImportDiagnostic(
                    "XLSB-FORMULA-PRESERVED",
                    XlsbImportDiagnosticSeverity.Warning,
                    $"Preserved an unsupported BIFF12 formula token stream at row {row}, column {column}; its cached value was projected.",
                    partName,
                    record.Type,
                    record.Offset));
                PreserveRecord(options, workbook, partName, record);
            }
            return cell;
        }

        private static byte[] CopyTail(byte[] data, int offset) {
            if (offset < 0 || offset > data.Length) throw new ArgumentOutOfRangeException(nameof(offset));
            byte[] result = new byte[data.Length - offset];
            Buffer.BlockCopy(data, offset, result, 0, result.Length);
            return result;
        }

        private static void PreserveRecord(XlsbImportOptions options, XlsbWorkbook workbook, string partName, XlsbRecord record) {
            if (!options.ReportPreservedRecords) return;
            workbook.AddPreservedRecord(new XlsbPreservedRecordInfo(partName, record.Type, record.Offset, record.Size));
        }
    }
}
