using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Rewrites a worksheet cell table while preserving all records outside it and unknown in-table metadata.</summary>
    internal static class XlsbWorksheetPartWriter {
        private const int BrtRowHdr = 0;
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
        private const int BrtBeginSheetData = 145;
        private const int BrtEndSheetData = 146;

        private static readonly byte[] DefaultRowPayload = {
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x40, 0x01, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x00, 0x00
        };

        internal static byte[] Rewrite(byte[] originalPart, IReadOnlyList<XlsbWriteCell> cells) {
            if (originalPart == null) throw new ArgumentNullException(nameof(originalPart));
            if (cells == null) throw new ArgumentNullException(nameof(cells));

            IReadOnlyList<XlsbRecord> records;
            using (var input = new MemoryStream(originalPart, writable: false)) {
                records = XlsbRecordReader.ReadAll(input);
            }

            int beginIndex = FindSingleRecord(records, BrtBeginSheetData, "BrtBeginSheetData");
            int endIndex = FindSingleRecord(records, BrtEndSheetData, "BrtEndSheetData");
            if (endIndex <= beginIndex) {
                throw new InvalidDataException("The XLSB worksheet has an invalid sheet-data boundary order.");
            }

            XlsbSheetDataLayout layout = ParseSheetDataLayout(records, beginIndex + 1, endIndex);
            IReadOnlyDictionary<int, IReadOnlyList<XlsbWriteCell>> cellsByRow = cells
                .GroupBy(cell => cell.Row - 1)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<XlsbWriteCell>)group.OrderBy(cell => cell.Column).ToArray());
            int[] rowIndexes = layout.Rows.Keys.Concat(cellsByRow.Keys).Distinct().OrderBy(row => row).ToArray();

            using var output = new MemoryStream(originalPart.Length + Math.Max(256, cells.Count * 24));
            for (int index = 0; index <= beginIndex; index++) {
                WriteRecord(output, records[index]);
            }

            foreach (XlsbRecord metadata in layout.PrefixRecords) {
                WriteRecord(output, metadata);
            }

            foreach (int rowIndex in rowIndexes) {
                if (layout.Rows.TryGetValue(rowIndex, out XlsbSourceRowBlock? sourceRow)) {
                    byte[] rowPayload = (byte[])sourceRow.RowHeader.Data.Clone();
                    WriteUInt32(rowPayload, 0, checked((uint)rowIndex));
                    XlsbRecordWriter.Write(output, BrtRowHdr, rowPayload);
                } else {
                    byte[] rowPayload = (byte[])DefaultRowPayload.Clone();
                    WriteUInt32(rowPayload, 0, checked((uint)rowIndex));
                    XlsbRecordWriter.Write(output, BrtRowHdr, rowPayload);
                }

                if (cellsByRow.TryGetValue(rowIndex, out IReadOnlyList<XlsbWriteCell>? rowCells)) {
                    foreach (XlsbWriteCell cell in rowCells) {
                        WriteCell(output, cell);
                    }
                }

                if (sourceRow != null) {
                    foreach (XlsbRecord metadata in sourceRow.PreservedRecords) {
                        WriteRecord(output, metadata);
                    }
                }
            }

            for (int index = endIndex; index < records.Count; index++) {
                WriteRecord(output, records[index]);
            }

            return output.ToArray();
        }

        private static XlsbSheetDataLayout ParseSheetDataLayout(IReadOnlyList<XlsbRecord> records, int start, int end) {
            var prefix = new List<XlsbRecord>();
            var rows = new Dictionary<int, XlsbSourceRowBlock>();
            XlsbSourceRowBlock? current = null;
            for (int index = start; index < end; index++) {
                XlsbRecord record = records[index];
                if (record.Type == BrtRowHdr) {
                    var cursor = new XlsbBinaryCursor(record.Data);
                    int rowIndex = cursor.ReadInt32();
                    if (rowIndex < 0 || rowIndex >= 1_048_576 || rows.ContainsKey(rowIndex)) {
                        throw new InvalidDataException($"The XLSB worksheet contains invalid or duplicate row index {rowIndex}.");
                    }

                    current = new XlsbSourceRowBlock(record);
                    rows.Add(rowIndex, current);
                } else if (!IsCellRecord(record.Type)) {
                    if (current == null) {
                        prefix.Add(record);
                    } else {
                        current.PreservedRecords.Add(record);
                    }
                }
            }

            return new XlsbSheetDataLayout(prefix, rows);
        }

        private static void WriteCell(Stream output, XlsbWriteCell cell) {
            if (cell.SourceRecordType.HasValue && cell.SourceRecordData != null) {
                XlsbRecordWriter.Write(output, cell.SourceRecordType.Value, cell.SourceRecordData);
                return;
            }

            using var payload = new MemoryStream();
            WriteUInt32(payload, checked((uint)(cell.Column - 1)));
            WriteUInt32(payload, cell.StyleIndex & 0x00FFFFFFU);
            int recordType;
            switch (cell.Kind) {
                case XlsbWriteCellKind.Blank:
                    recordType = BrtCellBlank;
                    break;
                case XlsbWriteCellKind.Number:
                    recordType = BrtCellReal;
                    WriteDouble(payload, Convert.ToDouble(cell.Value, System.Globalization.CultureInfo.InvariantCulture));
                    break;
                case XlsbWriteCellKind.Text:
                    recordType = BrtCellSt;
                    WriteWideString(payload, (string?)cell.Value ?? string.Empty);
                    break;
                case XlsbWriteCellKind.Boolean:
                    recordType = BrtCellBool;
                    payload.WriteByte((bool)cell.Value! ? (byte)1 : (byte)0);
                    break;
                case XlsbWriteCellKind.Error:
                    recordType = BrtCellError;
                    payload.WriteByte((byte)cell.Value!);
                    break;
                case XlsbWriteCellKind.FormulaNumber:
                    recordType = BrtFmlaNum;
                    WriteDouble(payload, Convert.ToDouble(cell.Value, System.Globalization.CultureInfo.InvariantCulture));
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaText:
                    recordType = BrtFmlaString;
                    WriteWideString(payload, (string?)cell.Value ?? string.Empty);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaBoolean:
                    recordType = BrtFmlaBool;
                    payload.WriteByte((bool)cell.Value! ? (byte)1 : (byte)0);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaError:
                    recordType = BrtFmlaError;
                    payload.WriteByte((byte)cell.Value!);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB write cell kind {cell.Kind}.");
            }

            XlsbRecordWriter.Write(output, recordType, payload.ToArray());
        }

        private static void WriteFormula(Stream payload, byte[]? formulaPayload) {
            byte[] bytes = formulaPayload ?? throw new InvalidOperationException("Formula cell has no preserved BIFF12 formula payload.");
            payload.Write(bytes, 0, bytes.Length);
        }

        private static void WriteWideString(Stream stream, string value) {
            WriteUInt32(stream, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteDouble(Stream stream, double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }

        private static void WriteUInt32(byte[] data, int offset, uint value) {
            if (offset < 0 || offset > data.Length - 4) throw new ArgumentOutOfRangeException(nameof(offset));
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteRecord(Stream stream, XlsbRecord record) =>
            XlsbRecordWriter.Write(stream, record.Type, record.Data);

        private static int FindSingleRecord(IReadOnlyList<XlsbRecord> records, int recordType, string recordName) {
            int found = -1;
            for (int index = 0; index < records.Count; index++) {
                if (records[index].Type != recordType) continue;
                if (found >= 0) throw new InvalidDataException($"The XLSB worksheet contains more than one {recordName} record.");
                found = index;
            }

            if (found < 0) throw new InvalidDataException($"The XLSB worksheet does not contain a {recordName} record.");
            return found;
        }

        private static bool IsCellRecord(int recordType) {
            return (recordType >= BrtCellBlank && recordType <= BrtFmlaError)
                || recordType == BrtCellRString;
        }

        private sealed class XlsbSheetDataLayout {
            internal XlsbSheetDataLayout(List<XlsbRecord> prefixRecords, Dictionary<int, XlsbSourceRowBlock> rows) {
                PrefixRecords = prefixRecords;
                Rows = rows;
            }

            internal IReadOnlyList<XlsbRecord> PrefixRecords { get; }

            internal IReadOnlyDictionary<int, XlsbSourceRowBlock> Rows { get; }
        }

        private sealed class XlsbSourceRowBlock {
            internal XlsbSourceRowBlock(XlsbRecord rowHeader) {
                RowHeader = rowHeader;
            }

            internal XlsbRecord RowHeader { get; }

            internal List<XlsbRecord> PreservedRecords { get; } = new List<XlsbRecord>();
        }
    }
}
