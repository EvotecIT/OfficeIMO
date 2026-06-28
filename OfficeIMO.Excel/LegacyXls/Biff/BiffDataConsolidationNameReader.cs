using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads DConName metadata from workbook globals.
    /// </summary>
    internal static class BiffDataConsolidationNameReader {
        internal static bool TryRead(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsDataConsolidationName? name) {
            name = null;
            if (record.Type != (ushort)BiffRecordType.DConName) {
                return false;
            }

            byte[] payload = record.Payload;
            if (payload.Length < 4) {
                AddInvalidDiagnostic(diagnostics, record, "The DConName record is too short.");
                return false;
            }

            try {
                int offset = 0;
                string definedName = BiffStringReader.ReadShortUnicodeString(payload, ref offset);
                if (string.IsNullOrWhiteSpace(definedName)) {
                    AddInvalidDiagnostic(diagnostics, record, "The DConName defined name is empty.");
                    return false;
                }

                if (offset + 2 > payload.Length) {
                    AddInvalidDiagnostic(diagnostics, record, "The DConName record ended before cchFile.");
                    return false;
                }

                ushort fileCharacterCount = BiffRecordReader.ReadUInt16(payload, offset);
                offset += 2;
                string source = string.Empty;
                LegacyXlsDataConsolidationSourceKind sourceKind = LegacyXlsDataConsolidationSourceKind.SelfReference;
                if (fileCharacterCount > 0) {
                    string rawSource = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, fileCharacterCount);
                    byte? sourcePrefix = rawSource.Length > 0 && rawSource[0] <= byte.MaxValue ? (byte)rawSource[0] : null;
                    sourceKind = sourcePrefix == 0x01
                        ? LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath
                        : sourcePrefix == 0x02
                            ? LegacyXlsDataConsolidationSourceKind.SelfReference
                            : LegacyXlsDataConsolidationSourceKind.Unknown;
                    source = sourceKind == LegacyXlsDataConsolidationSourceKind.Unknown
                        ? rawSource
                        : rawSource.Substring(1);
                }

                name = new LegacyXlsDataConsolidationName(
                    record.Offset,
                    record.Type,
                    definedName,
                    sourceKind,
                    source,
                    payload.Length - offset);
                return true;
            } catch (InvalidDataException ex) {
                AddInvalidDiagnostic(diagnostics, record, ex.Message);
                return false;
            } catch (OverflowException ex) {
                AddInvalidDiagnostic(diagnostics, record, ex.Message);
                return false;
            } catch (ArgumentOutOfRangeException ex) {
                AddInvalidDiagnostic(diagnostics, record, ex.Message);
                return false;
            }
        }

        private static void AddInvalidDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            BiffRecord record,
            string message) {
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                "XLS-BIFF-DCONNAME-INVALID",
                "The DConName named consolidation source record could not be decoded. " + message,
                recordOffset: record.Offset,
                recordType: record.Type,
                detailCode: "DConNameInvalid"));
        }
    }
}
