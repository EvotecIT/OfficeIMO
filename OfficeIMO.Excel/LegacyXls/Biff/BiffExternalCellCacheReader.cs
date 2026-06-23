using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads preserve-only external cell cache metadata from XCT and CRN record groups.
    /// </summary>
    internal static class BiffExternalCellCacheReader {
        internal static LegacyXlsExternalCellCache? ReadXct(
            BiffRecord record,
            LegacyXlsExternalReference? currentExternalReference,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (currentExternalReference == null) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XCT-ORPHANED",
                    "An XCT external cell cache record appeared before a SupBook supporting link.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return null;
            }

            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XCT-SHORT",
                    "An XCT external cell cache record was too short to parse.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return null;
            }

            int signedCrnCount = unchecked((short)BiffRecordReader.ReadUInt16(record.Payload, 0));
            int declaredCrnCount = signedCrnCount < 0 ? -signedCrnCount : signedCrnCount;
            int? sheetIndex = null;
            string? sheetName = null;
            if (record.Payload.Length >= 4) {
                ushort rawSheetIndex = BiffRecordReader.ReadUInt16(record.Payload, 2);
                sheetIndex = rawSheetIndex;
                if (rawSheetIndex < currentExternalReference.SheetNames.Count) {
                    sheetName = currentExternalReference.SheetNames[rawSheetIndex];
                } else {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-XCT-SHEET-INDEX-INVALID",
                        "An XCT external cell cache record references a sheet index outside the preceding SupBook sheet table.",
                        recordOffset: record.Offset,
                        recordType: record.Type,
                        detailCode: $"ExternalCacheSheet:{rawSheetIndex}"));
                }
            }

            var cache = new LegacyXlsExternalCellCache(declaredCrnCount, sheetIndex, sheetName, signedCrnCount >= 0);
            currentExternalReference.MutableCachedCellCaches.Add(cache);
            return cache;
        }

        internal static void ReadCrn(
            BiffRecord record,
            LegacyXlsExternalCellCache? currentCache,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (currentCache == null) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CRN-ORPHANED",
                    "A CRN external cell cache record appeared before an XCT cache section.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            try {
                ReadCrnCore(record, currentCache);
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CRN-INVALID",
                    $"A CRN external cell cache record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }
        }

        private static void ReadCrnCore(BiffRecord record, LegacyXlsExternalCellCache currentCache) {
            if (record.Payload.Length < 4) {
                throw new InvalidDataException("The CRN record was too short.");
            }

            int lastColumn = record.Payload[0];
            int firstColumn = record.Payload[1];
            if (lastColumn < firstColumn) {
                throw new InvalidDataException("The CRN last column is before its first column.");
            }

            int row = BiffRecordReader.ReadUInt16(record.Payload, 2);
            int offset = 4;
            int valueCount = checked(lastColumn - firstColumn + 1);
            for (int i = 0; i < valueCount; i++) {
                LegacyXlsExternalCachedCell cell = ReadCachedCell(record.Payload, ref offset, row, firstColumn + i);
                currentCache.MutableCells.Add(cell);
            }

            if (offset != record.Payload.Length) {
                throw new InvalidDataException("The CRN record contains trailing bytes after its cached values.");
            }
        }

        private static LegacyXlsExternalCachedCell ReadCachedCell(byte[] payload, ref int offset, int row, int column) {
            if (offset >= payload.Length) {
                throw new InvalidDataException("The CRN record ended before all cached values could be read.");
            }

            byte type = payload[offset];
            switch (type) {
                case 0x00:
                    EnsureAvailable(payload, offset, 9);
                    offset += 9;
                    return new LegacyXlsExternalCachedCell(row, column, LegacyXlsCellValueKind.Blank, null);
                case 0x01:
                    EnsureAvailable(payload, offset, 9);
                    double number = BiffRecordReader.ReadDouble(payload, offset + 1);
                    offset += 9;
                    return new LegacyXlsExternalCachedCell(row, column, LegacyXlsCellValueKind.Number, number);
                case 0x02:
                    offset++;
                    string text = BiffStringReader.ReadUnicodeString(payload, ref offset);
                    return new LegacyXlsExternalCachedCell(row, column, LegacyXlsCellValueKind.Text, text);
                case 0x04:
                    EnsureAvailable(payload, offset, 9);
                    bool boolean = payload[offset + 1] != 0;
                    offset += 9;
                    return new LegacyXlsExternalCachedCell(row, column, LegacyXlsCellValueKind.Boolean, boolean);
                case 0x10:
                    EnsureAvailable(payload, offset, 9);
                    string error = BiffErrorValue.ToText(payload[offset + 1]);
                    offset += 9;
                    return new LegacyXlsExternalCachedCell(row, column, LegacyXlsCellValueKind.Error, error);
                default:
                    throw new InvalidDataException($"Unsupported CRN cached value type 0x{type:X2}.");
            }
        }

        private static void EnsureAvailable(byte[] payload, int offset, int length) {
            if (offset + length > payload.Length) {
                throw new InvalidDataException("The CRN cached value ended early.");
            }
        }
    }
}
