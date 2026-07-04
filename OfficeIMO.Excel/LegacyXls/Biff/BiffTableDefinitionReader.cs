using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffTableDefinitionReader {
        private const ushort Feature11RecordType = 0x0872;
        private const ushort Feature12RecordType = 0x0878;
        private const int FeatureHeaderSize = 27;
        private const int List12HeaderSize = 18;
        private const int List12BlockLevelFixedSize = 36;
        private const int TableFeatureFixedSize = 64;
        private const ushort List12BlockLevel = 0x0000;
        private const ushort List12TableStyleClientInfo = 0x0001;
        private const ushort List12DisplayName = 0x0002;

        internal static bool IsTableDefinitionRecord(ushort recordType) {
            return recordType == (ushort)BiffRecordType.FeatHdr11
                || recordType == (ushort)BiffRecordType.Feature11
                || recordType == (ushort)BiffRecordType.List12
                || recordType == (ushort)BiffRecordType.Feature12;
        }

        internal static bool TryRead(
            BiffRecord record,
            LegacyXlsWorksheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out bool projectable) {
            projectable = false;

            if (record.Type == (ushort)BiffRecordType.FeatHdr11) {
                projectable = record.Payload.Length >= 12;
                if (!projectable) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-FEATHDR11-SHORT",
                        "The FeatHdr11 table-definition header is shorter than expected.",
                        sheetName: sheet.Name,
                        recordOffset: record.Offset,
                        recordType: record.Type));
                }

                return true;
            }

            if (record.Type == (ushort)BiffRecordType.List12) {
                if (TryReadList12(record, sheet, diagnostics)) {
                    projectable = true;
                }

                return true;
            }

            if (record.Type != (ushort)BiffRecordType.Feature11
                && record.Type != (ushort)BiffRecordType.Feature12) {
                return IsTableDefinitionRecord(record.Type);
            }

            if (TryReadTableFeature(record, sheet.Name, diagnostics, out LegacyXlsTableDefinition? tableDefinition)) {
                sheet.AddTableDefinition(tableDefinition!);
                projectable = true;
            }

            return true;
        }

        private static bool TryReadList12(
            BiffRecord record,
            LegacyXlsWorksheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < List12HeaderSize) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-SHORT", "The List12 table-definition record is shorter than expected.");
                return false;
            }

            ushort embeddedRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (embeddedRecordType != record.Type) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-HEADER-MISMATCH", "The List12 record has an unexpected future-record header.");
                return false;
            }

            ushort listDataType = BiffRecordReader.ReadUInt16(record.Payload, 12);
            uint idList = BiffRecordReader.ReadUInt32(record.Payload, 14);
            return listDataType switch {
                List12BlockLevel => TryReadList12BlockLevel(record, sheet, diagnostics, idList),
                List12TableStyleClientInfo => TryReadList12TableStyle(record, sheet, diagnostics, idList),
                List12DisplayName => TryReadList12DisplayName(record, sheet, diagnostics, idList),
                _ => false
            };
        }

        private static bool TryReadList12BlockLevel(
            BiffRecord record,
            LegacyXlsWorksheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            uint idList) {
            if (record.Payload.Length < List12HeaderSize + List12BlockLevelFixedSize) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-SHORT", "The List12 block-level formatting record is shorter than expected.");
                return false;
            }

            int offset = List12HeaderSize;
            int headerDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset);
            int headerStyleRecordIndex = BiffRecordReader.ReadInt32(record.Payload, offset + 4);
            int dataDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset + 8);
            int dataStyleRecordIndex = BiffRecordReader.ReadInt32(record.Payload, offset + 12);
            int totalDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset + 16);
            int totalStyleRecordIndex = BiffRecordReader.ReadInt32(record.Payload, offset + 20);
            int borderDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset + 24);
            int headerBorderDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset + 28);
            int totalBorderDxfByteCount = BiffRecordReader.ReadInt32(record.Payload, offset + 32);
            int[] byteCounts = {
                headerDxfByteCount,
                dataDxfByteCount,
                totalDxfByteCount,
                borderDxfByteCount,
                headerBorderDxfByteCount,
                totalBorderDxfByteCount
            };
            if (byteCounts.Any(count => count < 0)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-SIZE-INVALID", "The List12 block-level formatting record contains an invalid negative DXF byte count.");
                return false;
            }

            if (byteCounts.Any(count => count > 0)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-DXF-UNSUPPORTED", "List12 block-level DXF payloads are not yet projected.");
                return false;
            }

            offset += List12BlockLevelFixedSize;
            if (!TryReadOptionalStyleName(record, sheet.Name, diagnostics, headerStyleRecordIndex, ref offset, "header", out string? headerStyleName)
                || !TryReadOptionalStyleName(record, sheet.Name, diagnostics, dataStyleRecordIndex, ref offset, "data", out string? dataStyleName)
                || !TryReadOptionalStyleName(record, sheet.Name, diagnostics, totalStyleRecordIndex, ref offset, "total-row", out string? totalStyleName)) {
                return false;
            }

            var formatting = new LegacyXlsTableBlockLevelFormatting(
                ToOptionalStyleRecordIndex(headerStyleRecordIndex),
                headerStyleName,
                ToOptionalStyleRecordIndex(dataStyleRecordIndex),
                dataStyleName,
                ToOptionalStyleRecordIndex(totalStyleRecordIndex),
                totalStyleName);
            if (!sheet.TryApplyTableDefinitionBlockLevelFormatting(idList, formatting)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-TABLE-MISSING", "The List12 block-level formatting record did not match a projected table definition.");
                return false;
            }

            return true;
        }

        private static bool TryReadList12TableStyle(
            BiffRecord record,
            LegacyXlsWorksheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            uint idList) {
            if (record.Payload.Length < List12HeaderSize + 3) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-STYLE-SHORT", "The List12 table-style record is shorter than expected.");
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, List12HeaderSize);
            int styleNameOffset = List12HeaderSize + 2;
            string styleName;
            try {
                styleName = BiffStringReader.ReadUnicodeString(record.Payload, ref styleNameOffset);
            } catch (InvalidDataException ex) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-STYLE-INVALID", "The List12 table-style name could not be read. " + ex.Message);
                return false;
            }

            if (string.IsNullOrWhiteSpace(styleName)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-STYLE-MISSING", "The List12 table-style record does not contain a usable style name.");
                return false;
            }

            if (!sheet.TryApplyTableDefinitionStyle(
                idList,
                styleName,
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-TABLE-MISSING", "The List12 table-style record did not match a projected table definition.");
                return false;
            }

            return true;
        }

        private static bool TryReadList12DisplayName(
            BiffRecord record,
            LegacyXlsWorksheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            uint idList) {
            if (record.Payload.Length < List12HeaderSize + 6) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-DISPLAY-NAME-SHORT", "The List12 display-name record is shorter than expected.");
                return false;
            }

            string displayName;
            string comment;
            int offset = List12HeaderSize;
            try {
                displayName = BiffStringReader.ReadUnicodeString(record.Payload, ref offset);
                comment = BiffStringReader.ReadUnicodeString(record.Payload, ref offset);
            } catch (InvalidDataException ex) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-DISPLAY-NAME-INVALID", "The List12 display-name record could not be read. " + ex.Message);
                return false;
            }

            if (string.IsNullOrWhiteSpace(displayName) && string.IsNullOrWhiteSpace(comment)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-DISPLAY-NAME-MISSING", "The List12 display-name record does not contain a usable display name or comment.");
                return false;
            }

            if (!sheet.TryApplyTableDefinitionDisplayMetadata(idList, displayName, comment)) {
                AddWarning(record, sheet.Name, diagnostics, "XLS-BIFF-LIST12-DISPLAY-NAME-TABLE-MISSING", "The List12 display-name record did not match a projected table definition.");
                return false;
            }

            return true;
        }

        private static bool TryReadOptionalStyleName(
            BiffRecord record,
            string sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            int styleRecordIndex,
            ref int offset,
            string regionName,
            out string? styleName) {
            styleName = null;
            if (styleRecordIndex == -1) {
                return true;
            }

            if (styleRecordIndex < -1) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-STYLE-INDEX-INVALID", $"The List12 block-level {regionName} style index is invalid.");
                return false;
            }

            try {
                styleName = BiffStringReader.ReadUnicodeString(record.Payload, ref offset);
                return true;
            } catch (InvalidDataException ex) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-LIST12-BLOCK-LEVEL-STYLE-NAME-INVALID", $"The List12 block-level {regionName} style name could not be read. {ex.Message}");
                return false;
            }
        }

        private static int? ToOptionalStyleRecordIndex(int styleRecordIndex) {
            return styleRecordIndex >= 0 ? styleRecordIndex : null;
        }

        private static bool TryReadTableFeature(
            BiffRecord record,
            string sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsTableDefinition? tableDefinition) {
            tableDefinition = null;
            if (record.Payload.Length < FeatureHeaderSize + 8 + TableFeatureFixedSize + 3) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-SHORT", "The table definition record is shorter than the supported Feature11/Feature12 shape.");
                return false;
            }

            ushort embeddedRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (embeddedRecordType != record.Type
                || (embeddedRecordType != Feature11RecordType && embeddedRecordType != Feature12RecordType)) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-HEADER-MISMATCH", "The table definition record has an unexpected future-record header.");
                return false;
            }

            ushort referenceCount = BiffRecordReader.ReadUInt16(record.Payload, 19);
            if (referenceCount != 1) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-RANGE-COUNT", "Only single-range legacy table definitions are supported for projection.");
                return false;
            }

            int tableFeatureOffset = FeatureHeaderSize + 8;
            uint sourceType = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset);
            uint idList = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset + 4);
            uint headerRows = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset + 8);
            uint totalRows = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset + 12);
            uint fixedSize = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset + 20);
            uint flags = BiffRecordReader.ReadUInt32(record.Payload, tableFeatureOffset + 28);
            if (sourceType != 0 || idList == 0 || fixedSize != TableFeatureFixedSize || headerRows > 1 || totalRows > 1) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-SUBSET", "Only local worksheet tables with optional header row and up to one totals row are supported for projection.");
                return false;
            }

            int stringOffset = tableFeatureOffset + TableFeatureFixedSize;
            string tableName = BiffStringReader.ReadUnicodeString(record.Payload, ref stringOffset);
            if (string.IsNullOrWhiteSpace(tableName)) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-NAME-MISSING", "The table definition does not contain a usable table name.");
                return false;
            }

            if (stringOffset + 2 > record.Payload.Length) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-COLUMNS-MISSING", "The table definition does not contain the declared field count.");
                return false;
            }

            ushort fieldCount = BiffRecordReader.ReadUInt16(record.Payload, stringOffset);
            string range = ReadReference(record.Payload, FeatureHeaderSize);
            if (!TryGetColumnCount(range, out int rangeColumnCount) || fieldCount > 0 && fieldCount != rangeColumnCount) {
                AddWarning(record, sheetName, diagnostics, "XLS-BIFF-TABLE-DEFINITION-COLUMN-MISMATCH", "The table definition field count does not match its range.");
                return false;
            }

            tableDefinition = new LegacyXlsTableDefinition(
                tableName,
                range,
                headerRows == 1,
                totalRows,
                (flags & 0x00000002) != 0,
                idList,
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static string ReadReference(byte[] payload, int offset) {
            ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
            return firstRow == lastRow && firstColumn == lastColumn
                ? A1.CellReference(firstRow + 1, firstColumn + 1)
                : $"{A1.CellReference(firstRow + 1, firstColumn + 1)}:{A1.CellReference(lastRow + 1, lastColumn + 1)}";
        }

        private static bool TryGetColumnCount(string range, out int columnCount) {
            columnCount = 0;
            if (!A1.TryParseStrictRange(range, out _, out int startColumn, out _, out int endColumn)) {
                return false;
            }

            columnCount = endColumn - startColumn + 1;
            return columnCount > 0;
        }

        private static void AddWarning(
            BiffRecord record,
            string sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            string code,
            string message) {
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                code,
                message,
                sheetName: sheetName,
                recordOffset: record.Offset,
                recordType: record.Type));
        }
    }
}
