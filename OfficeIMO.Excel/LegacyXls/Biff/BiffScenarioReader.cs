using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffScenarioReader {
        private const ushort ScenManRecordType = 0x00AE;
        private const ushort ScenarioRecordType = 0x00AF;

        internal static bool TryReadManager(BiffRecord record, out LegacyXlsScenarioManager? manager) {
            manager = null;
            byte[] payload = record.Payload;
            if (record.Type != ScenManRecordType || payload.Length < 8) {
                return false;
            }

            ushort scenarioCount = BiffRecordReader.ReadUInt16(payload, 0);
            short currentScenarioIndex = BiffRecordReader.ReadInt16(payload, 2);
            short shownScenarioIndex = BiffRecordReader.ReadInt16(payload, 4);
            ushort resultRangeCount = BiffRecordReader.ReadUInt16(payload, 6);
            int expectedLength = checked(8 + (resultRangeCount * 8));
            if (resultRangeCount > 32 || payload.Length != expectedLength) {
                return false;
            }

            var ranges = new List<string>(resultRangeCount);
            for (int i = 0; i < resultRangeCount; i++) {
                if (!TryReadCellRange(payload, 8 + (i * 8), out string? reference)) {
                    return false;
                }

                ranges.Add(reference!);
            }

            manager = new LegacyXlsScenarioManager(scenarioCount, currentScenarioIndex, shownScenarioIndex, ranges);
            return true;
        }

        internal static bool TryReadScenario(BiffRecord record, out LegacyXlsScenario? scenario) {
            scenario = null;
            byte[] payload = record.Payload;
            if (record.Type != ScenarioRecordType || payload.Length < 8) {
                return false;
            }

            ushort inputCellCount = BiffRecordReader.ReadUInt16(payload, 0);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 2);
            byte nameLength = payload[4];
            byte userLength = payload[5];
            byte commentLength = payload[6];
            if (inputCellCount == 0 || inputCellCount > 32 || nameLength == 0 || (flags & 0xfffc) != 0) {
                return false;
            }

            int offset = 8;
            int referencesLength = checked(inputCellCount * 4);
            if (offset + referencesLength > payload.Length) {
                return false;
            }

            var references = new List<ScenarioCellReference>(inputCellCount);
            for (int i = 0; i < inputCellCount; i++) {
                if (!TryReadCellReference(payload, offset + (i * 4), out ScenarioCellReference cellReference)) {
                    return false;
                }

                references.Add(cellReference);
            }

            offset += referencesLength;
            if (!TryReadUnicodeStringNoCch(payload, ref offset, nameLength, out string? name)) {
                return false;
            }

            if (!TryReadOptionalUnicodeString(payload, ref offset, userLength, out string? user)
                || !TryReadOptionalUnicodeString(payload, ref offset, commentLength, out string? comment)) {
                return false;
            }

            var inputCells = new List<LegacyXlsScenarioInputCell>(inputCellCount);
            for (int i = 0; i < inputCellCount; i++) {
                if (!TryReadUnicodeString(payload, ref offset, out string? value)) {
                    return false;
                }

                ScenarioCellReference reference = references[i];
                inputCells.Add(new LegacyXlsScenarioInputCell(reference.CellReference, reference.Row, reference.Column, reference.Deleted, value!));
            }

            if (offset != payload.Length) {
                return false;
            }

            scenario = new LegacyXlsScenario(
                name!,
                locked: (flags & 0x0001) != 0,
                hidden: (flags & 0x0002) != 0,
                user,
                comment,
                inputCells);
            return true;
        }

        private static bool TryReadCellReference(byte[] payload, int offset, out ScenarioCellReference reference) {
            reference = default;
            if (offset + 4 > payload.Length) {
                return false;
            }

            ushort row = BiffRecordReader.ReadUInt16(payload, offset);
            ushort columnBits = BiffRecordReader.ReadUInt16(payload, offset + 2);
            int column = columnBits & 0x00ff;
            if (column > 0x00ff) {
                return false;
            }

            int oneBasedRow = row + 1;
            int oneBasedColumn = column + 1;
            reference = new ScenarioCellReference(
                A1.CellReference(oneBasedRow, oneBasedColumn),
                oneBasedRow,
                oneBasedColumn,
                (columnBits & 0x4000) != 0);
            return true;
        }

        private static bool TryReadCellRange(byte[] payload, int offset, out string? reference) {
            reference = null;
            if (offset + 8 > payload.Length) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
            if (lastRow < firstRow || lastColumn < firstColumn || firstColumn > 0x00ff || lastColumn > 0x00ff) {
                return false;
            }

            string start = A1.CellReference(firstRow + 1, firstColumn + 1);
            string end = A1.CellReference(lastRow + 1, lastColumn + 1);
            reference = start == end ? start : start + ":" + end;
            return true;
        }

        private static bool TryReadOptionalUnicodeString(byte[] payload, ref int offset, int charCount, out string? value) {
            value = null;
            if (charCount == 0) {
                return true;
            }

            return TryReadUnicodeStringNoCch(payload, ref offset, charCount, out value);
        }

        private static bool TryReadUnicodeString(byte[] payload, ref int offset, out string? value) {
            value = null;
            try {
                value = BiffStringReader.ReadUnicodeString(payload, ref offset);
                return true;
            } catch (InvalidDataException) {
                return false;
            }
        }

        private static bool TryReadUnicodeStringNoCch(byte[] payload, ref int offset, int charCount, out string? value) {
            value = null;
            try {
                value = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, charCount);
                return true;
            } catch (InvalidDataException) {
                return false;
            }
        }

        private readonly struct ScenarioCellReference {
            internal ScenarioCellReference(string cellReference, int row, int column, bool deleted) {
                CellReference = cellReference;
                Row = row;
                Column = column;
                Deleted = deleted;
            }

            internal string CellReference { get; }

            internal int Row { get; }

            internal int Column { get; }

            internal bool Deleted { get; }
        }
    }
}
