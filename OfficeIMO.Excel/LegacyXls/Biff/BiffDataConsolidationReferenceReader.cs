using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads preserve-only DConRef metadata from workbook globals.
    /// </summary>
    internal static class BiffDataConsolidationReferenceReader {
        internal static bool TryRead(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsDataConsolidationReference? reference) {
            reference = null;
            if (record.Type != (ushort)BiffRecordType.DConRef) {
                return false;
            }

            byte[] payload = record.Payload;
            if (payload.Length < 8) {
                AddInvalidDiagnostic(diagnostics, record, "The DConRef record is too short.");
                return false;
            }

            try {
                ushort firstRow = BiffRecordReader.ReadUInt16(payload, 0);
                ushort lastRow = BiffRecordReader.ReadUInt16(payload, 2);
                byte firstColumn = payload[4];
                byte lastColumn = payload[5];
                ushort sourceCharacterCount = BiffRecordReader.ReadUInt16(payload, 6);
                int offset = 8;
                if (sourceCharacterCount < 2) {
                    AddInvalidDiagnostic(diagnostics, record, "The DConRef DConFile string is shorter than the minimum two characters.");
                    return false;
                }

                string rawSource = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, sourceCharacterCount);
                if (rawSource.Length == 0) {
                    AddInvalidDiagnostic(diagnostics, record, "The DConRef DConFile string is empty.");
                    return false;
                }

                byte? sourcePrefix = rawSource[0] <= byte.MaxValue ? (byte)rawSource[0] : null;
                LegacyXlsDataConsolidationSourceKind sourceKind = sourcePrefix == 0x01
                    ? LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath
                    : sourcePrefix == 0x02
                        ? LegacyXlsDataConsolidationSourceKind.SelfReference
                        : LegacyXlsDataConsolidationSourceKind.Unknown;
                string source = sourceKind == LegacyXlsDataConsolidationSourceKind.Unknown
                    ? rawSource
                    : rawSource.Substring(1);

                int firstRowOneBased = checked(firstRow + 1);
                int lastRowOneBased = checked(lastRow + 1);
                int firstColumnOneBased = checked(firstColumn + 1);
                int lastColumnOneBased = checked(lastColumn + 1);
                string start = A1.CellReference(firstRowOneBased, firstColumnOneBased);
                string end = A1.CellReference(lastRowOneBased, lastColumnOneBased);
                string cellRange = string.Equals(start, end, StringComparison.Ordinal)
                    ? start
                    : start + ":" + end;

                reference = new LegacyXlsDataConsolidationReference(
                    record.Offset,
                    record.Type,
                    firstRowOneBased,
                    lastRowOneBased,
                    firstColumnOneBased,
                    lastColumnOneBased,
                    cellRange,
                    sourceKind,
                    source,
                    sourcePrefix,
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
                "XLS-BIFF-DCONREF-INVALID",
                "The DConRef external data source record could not be decoded. " + message,
                recordOffset: record.Offset,
                recordType: record.Type,
                detailCode: "DConRefInvalid"));
        }
    }
}
