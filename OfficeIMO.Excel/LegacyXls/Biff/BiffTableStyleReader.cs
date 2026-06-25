using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffTableStyleReader {
        internal static bool TryRead(BiffRecord record, LegacyXlsWorkbook workbook, List<LegacyXlsImportDiagnostic> diagnostics) {
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.TableStyles:
                    if (TryReadTableStyles(record, diagnostics, out LegacyXlsTableStyleCollection? collection)) {
                        workbook.MutableTableStyleCollections.Add(collection!);
                    }

                    return true;

                case BiffRecordType.TableStyle:
                    if (TryReadTableStyle(record, diagnostics, out LegacyXlsTableStyle? style)) {
                        workbook.MutableTableStyles.Add(style!);
                    }

                    return true;

                case BiffRecordType.TableStyleElement:
                    if (TryReadTableStyleElement(record, diagnostics, out LegacyXlsTableStyleElement? element)) {
                        workbook.AddTableStyleElement(element!);
                    }

                    return true;

                default:
                    return false;
            }
        }

        private static bool TryReadTableStyles(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsTableStyleCollection? collection) {
            collection = null;
            if (record.Payload.Length < 20) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-TABLESTYLES-SHORT",
                    "The TableStyles record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }

            ushort cchDefaultTableStyle = BiffRecordReader.ReadUInt16(record.Payload, 16);
            ushort cchDefaultPivotStyle = BiffRecordReader.ReadUInt16(record.Payload, 18);
            int offset = 20;
            string? defaultTableStyleName = ReadUnicodeCharacters(record, diagnostics, ref offset, cchDefaultTableStyle, "default table style");
            string? defaultPivotStyleName = ReadUnicodeCharacters(record, diagnostics, ref offset, cchDefaultPivotStyle, "default PivotTable style");

            collection = new LegacyXlsTableStyleCollection(
                BiffRecordReader.ReadUInt32(record.Payload, 12),
                defaultTableStyleName,
                defaultPivotStyleName,
                BiffRecordReader.ReadUInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static bool TryReadTableStyle(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsTableStyle? style) {
            style = null;
            if (record.Payload.Length < 20) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-TABLESTYLE-SHORT",
                    "The TableStyle record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 12);
            ushort cchName = BiffRecordReader.ReadUInt16(record.Payload, 18);
            int offset = 20;
            string? name = ReadUnicodeCharacters(record, diagnostics, ref offset, cchName, "table style name");
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            style = new LegacyXlsTableStyle(
                name,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                BiffRecordReader.ReadUInt32(record.Payload, 14),
                BiffRecordReader.ReadUInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static bool TryReadTableStyleElement(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsTableStyleElement? element) {
            element = null;
            if (record.Payload.Length < 24) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-TABLESTYLEELEMENT-SHORT",
                    "The TableStyleElement record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }

            uint elementType = BiffRecordReader.ReadUInt32(record.Payload, 12);
            element = new LegacyXlsTableStyleElement(
                elementType,
                GetElementTypeName(elementType),
                BiffRecordReader.ReadUInt32(record.Payload, 16),
                BiffRecordReader.ReadUInt32(record.Payload, 20),
                BiffRecordReader.ReadUInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static string? ReadUnicodeCharacters(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            ref int offset,
            ushort characterCount,
            string fieldName) {
            if (characterCount == 0) {
                return null;
            }

            int byteCount = checked(characterCount * 2);
            if (offset + byteCount > record.Payload.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-TABLESTYLE-STRING-TRUNCATED",
                    $"The {fieldName} extends beyond the end of the table style record.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                offset = record.Payload.Length;
                return null;
            }

            string value = System.Text.Encoding.Unicode.GetString(record.Payload, offset, byteCount);
            offset += byteCount;
            return value;
        }

        private static string GetElementTypeName(uint elementType) {
            return elementType switch {
                0x00000000 => "WholeTable",
                0x00000001 => "HeaderRow",
                0x00000002 => "TotalRow",
                0x00000003 => "FirstColumn",
                0x00000004 => "LastColumn",
                0x00000005 => "RowStripe1",
                0x00000006 => "RowStripe2",
                0x00000007 => "ColumnStripe1",
                0x00000008 => "ColumnStripe2",
                0x00000009 => "FirstHeaderCell",
                0x0000000A => "LastHeaderCell",
                0x0000000B => "FirstTotalCell",
                0x0000000C => "LastTotalCell",
                0x0000000D => "FirstSubtotalColumn",
                0x0000000E => "SecondSubtotalColumn",
                0x0000000F => "ThirdSubtotalColumn",
                0x00000010 => "FirstSubtotalRow",
                0x00000011 => "SecondSubtotalRow",
                0x00000012 => "ThirdSubtotalRow",
                0x00000013 => "BlankRow",
                0x00000014 => "FirstColumnSubheading",
                0x00000015 => "SecondColumnSubheading",
                0x00000016 => "ThirdColumnSubheading",
                0x00000017 => "FirstRowSubheading",
                0x00000018 => "SecondRowSubheading",
                0x00000019 => "ThirdRowSubheading",
                0x0000001A => "PageFieldLabels",
                0x0000001B => "PageFieldValues",
                _ => $"Unknown:0x{elementType:X8}"
            };
        }
    }
}
