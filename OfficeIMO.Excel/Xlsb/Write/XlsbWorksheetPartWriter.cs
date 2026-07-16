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
        private const int BrtWsDim = 148;

        private static readonly byte[] DefaultRowProperties = {
            0x00, 0x00, 0x00, 0x00,
            0x2C, 0x01,
            0x00, 0x00, 0x00
        };

        internal static byte[] Create(
            ExcelSheet sheet,
            IReadOnlyList<XlsbWriteCell> cells,
            int cellFormatCount) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (cells == null) throw new ArgumentNullException(nameof(cells));

            XlsbWorksheetGeometryPlan geometry = XlsbWorksheetGeometryPlan.Create(sheet, cells, cellFormatCount);

            IReadOnlyDictionary<int, IReadOnlyList<XlsbWriteCell>> cellsByRow = cells
                .GroupBy(cell => cell.Row - 1)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<XlsbWriteCell>)group.OrderBy(cell => cell.Column).ToArray());
            int[] rowIndexes = cellsByRow.Keys.Concat(geometry.RowProperties.Keys).Distinct().OrderBy(row => row).ToArray();

            using var output = new MemoryStream(Math.Max(256, cells.Count * 24));
            XlsbRecordWriter.Write(output, 129); // BrtBeginSheet
            XlsbRecordWriter.Write(output, BrtWsDim, geometry.DimensionPayload);
            foreach (XlsbGeneratedRecord record in geometry.PrefixRecords) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
            XlsbRecordWriter.Write(output, BrtBeginSheetData);
            foreach (int rowIndex in rowIndexes) {
                cellsByRow.TryGetValue(rowIndex, out IReadOnlyList<XlsbWriteCell>? rowCells);
                geometry.RowProperties.TryGetValue(rowIndex, out byte[]? rowProperties);
                XlsbRecordWriter.Write(output, BrtRowHdr, CreateRowHeaderPayload(
                    rowIndex,
                    sourcePayload: null,
                    rowCells ?? Array.Empty<XlsbWriteCell>(),
                    rowProperties));
                if (rowCells == null) continue;
                foreach (XlsbWriteCell cell in rowCells) {
                    WriteCell(output, cell);
                }
            }
            XlsbRecordWriter.Write(output, BrtEndSheetData);
            foreach (XlsbGeneratedRecord record in geometry.SuffixRecords) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
            XlsbRecordWriter.Write(output, 130); // BrtEndSheet
            return output.ToArray();
        }

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
            byte[]? dimensionPayload = CreateExpandedDimensionPayload(records, cells);

            using var output = new MemoryStream(originalPart.Length + Math.Max(256, cells.Count * 24));
            for (int index = 0; index <= beginIndex; index++) {
                if (records[index].Type == BrtWsDim && dimensionPayload != null) {
                    XlsbRecordWriter.Write(output, BrtWsDim, dimensionPayload);
                } else {
                    WriteRecord(output, records[index]);
                }
            }

            foreach (XlsbRecord metadata in layout.PrefixRecords) {
                WriteRecord(output, metadata);
            }

            foreach (int rowIndex in rowIndexes) {
                layout.Rows.TryGetValue(rowIndex, out XlsbSourceRowBlock? sourceRow);
                cellsByRow.TryGetValue(rowIndex, out IReadOnlyList<XlsbWriteCell>? rowCells);
                byte[] rowPayload = CreateRowHeaderPayload(rowIndex, sourceRow?.RowHeader.Data, rowCells ?? Array.Empty<XlsbWriteCell>(), newProperties: null);
                XlsbRecordWriter.Write(output, BrtRowHdr, rowPayload);

                if (rowCells != null) {
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

        private static byte[] CreateRowHeaderPayload(
            int zeroBasedRow,
            byte[]? sourcePayload,
            IReadOnlyList<XlsbWriteCell> cells,
            byte[]? newProperties) {
            if (sourcePayload != null && sourcePayload.Length < 17) {
                throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} is truncated.");
            }

            var spans = ReadSourceSpans(sourcePayload, zeroBasedRow);
            foreach (IGrouping<int, XlsbWriteCell> group in cells.GroupBy(cell => (cell.Column - 1) / 1024)) {
                uint first = checked((uint)(group.Min(cell => cell.Column) - 1));
                uint last = checked((uint)(group.Max(cell => cell.Column) - 1));
                if (spans.TryGetValue(group.Key, out (uint First, uint Last) sourceSpan)) {
                    first = Math.Min(first, sourceSpan.First);
                    last = Math.Max(last, sourceSpan.Last);
                }
                spans[group.Key] = (first, last);
            }
            if (spans.Count > 16) {
                throw new InvalidDataException($"The XLSB row {zeroBasedRow + 1} requires {spans.Count} column spans, exceeding the BIFF12 limit of 16.");
            }

            using var payload = new MemoryStream(17 + spans.Count * 8);
            WriteUInt32(payload, checked((uint)zeroBasedRow));
            if (sourcePayload != null) {
                payload.Write(sourcePayload, 4, 9);
            } else if (newProperties != null) {
                if (newProperties.Length != 9) throw new InvalidDataException("A generated XLSB row-property payload must contain 9 bytes.");
                payload.Write(newProperties, 0, newProperties.Length);
            } else {
                payload.Write(DefaultRowProperties, 0, DefaultRowProperties.Length);
            }
            WriteUInt32(payload, checked((uint)spans.Count));
            foreach (KeyValuePair<int, (uint First, uint Last)> span in spans.OrderBy(pair => pair.Key)) {
                WriteUInt32(payload, span.Value.First);
                WriteUInt32(payload, span.Value.Last);
            }
            return payload.ToArray();
        }

        private static Dictionary<int, (uint First, uint Last)> ReadSourceSpans(byte[]? sourcePayload, int zeroBasedRow) {
            var spans = new Dictionary<int, (uint First, uint Last)>();
            if (sourcePayload == null) return spans;

            var cursor = new XlsbBinaryCursor(sourcePayload);
            cursor.Skip(13);
            uint count = cursor.ReadUInt32();
            if (count > 16 || cursor.Remaining != checked((int)count * 8)) {
                throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} has an invalid column-span payload.");
            }
            for (uint index = 0; index < count; index++) {
                uint first = cursor.ReadUInt32();
                uint last = cursor.ReadUInt32();
                int segment = checked((int)(first / 1024U));
                if (first > last || last >= 16_384U || first / 1024U != last / 1024U || spans.ContainsKey(segment)) {
                    throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} contains an invalid column span.");
                }
                spans.Add(segment, (first, last));
            }
            return spans;
        }

        private static byte[]? CreateExpandedDimensionPayload(
            IReadOnlyList<XlsbRecord> records,
            IReadOnlyList<XlsbWriteCell> cells) {
            XlsbRecord? dimension = null;
            foreach (XlsbRecord record in records) {
                if (record.Type != BrtWsDim) continue;
                if (dimension != null) {
                    throw new InvalidDataException("The XLSB worksheet contains more than one BrtWsDim record.");
                }
                dimension = record;
            }

            if (dimension == null) return null;
            if (dimension.Data.Length != 16) {
                throw new InvalidDataException($"The XLSB BrtWsDim record has invalid payload length {dimension.Data.Length}.");
            }
            if (cells.Count == 0) return (byte[])dimension.Data.Clone();

            var cursor = new XlsbBinaryCursor(dimension.Data);
            uint firstRow = cursor.ReadUInt32();
            uint lastRow = cursor.ReadUInt32();
            uint firstColumn = cursor.ReadUInt32();
            uint lastColumn = cursor.ReadUInt32();
            uint cellFirstRow = checked((uint)(cells.Min(cell => cell.Row) - 1));
            uint cellLastRow = checked((uint)(cells.Max(cell => cell.Row) - 1));
            uint cellFirstColumn = checked((uint)(cells.Min(cell => cell.Column) - 1));
            uint cellLastColumn = checked((uint)(cells.Max(cell => cell.Column) - 1));
            bool hasSourceCells = records.Any(record => IsCellRecord(record.Type));

            using var payload = new MemoryStream(16);
            WriteUInt32(payload, hasSourceCells ? Math.Min(firstRow, cellFirstRow) : cellFirstRow);
            WriteUInt32(payload, hasSourceCells ? Math.Max(lastRow, cellLastRow) : cellLastRow);
            WriteUInt32(payload, hasSourceCells ? Math.Min(firstColumn, cellFirstColumn) : cellFirstColumn);
            WriteUInt32(payload, hasSourceCells ? Math.Max(lastColumn, cellLastColumn) : cellLastColumn);
            return payload.ToArray();
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
