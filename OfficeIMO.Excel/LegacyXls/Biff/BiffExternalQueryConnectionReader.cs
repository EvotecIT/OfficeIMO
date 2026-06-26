using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads preserve-only DBQueryExt metadata from query table and PivotCache streams.
    /// </summary>
    internal static class BiffExternalQueryConnectionReader {
        private const int FixedPayloadLength = 28;
        private const int ParameterFlagByteLength = 2;

        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsExternalQueryConnection? connection) {
            connection = null;
            if (record.Type != (ushort)BiffRecordType.DbQueryExt) {
                return false;
            }

            byte[] payload = record.Payload;
            if (payload.Length < 4) {
                return false;
            }

            ushort futureRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            if (futureRecordType != 0x0803) {
                return false;
            }

            if (payload.Length < FixedPayloadLength) {
                AddInvalidDiagnostic(diagnostics, record, "The DBQueryExt record is shorter than the fixed header.");
                return false;
            }

            ushort dataSourceType = BiffRecordReader.ReadUInt16(payload, 4);
            ushort connectionFlags = BiffRecordReader.ReadUInt16(payload, 6);
            ushort sourceSpecificFlags = BiffRecordReader.ReadUInt16(payload, 8);
            ushort queryOptions = BiffRecordReader.ReadUInt16(payload, 10);
            byte editVersion = payload[12];
            byte refreshedVersion = payload[13];
            byte refreshableMinimumVersion = payload[14];
            ushort oleDbConnectionCount = BiffRecordReader.ReadUInt16(payload, 18);
            ushort futureByteCount = BiffRecordReader.ReadUInt16(payload, 20);
            ushort refreshIntervalMinutes = BiffRecordReader.ReadUInt16(payload, 22);
            ushort htmlFormat = BiffRecordReader.ReadUInt16(payload, 24);
            ushort parameterFlagCount = BiffRecordReader.ReadUInt16(payload, 26);
            int parameterFlagByteCount = payload.Length - FixedPayloadLength - futureByteCount;
            if (parameterFlagByteCount < 0) {
                AddInvalidDiagnostic(diagnostics, record, "The DBQueryExt future byte count exceeds the remaining payload.");
                return false;
            }

            LegacyXlsExternalQueryConnectionSourceType sourceTypeKind = ToSourceTypeKind(dataSourceType);
            connection = new LegacyXlsExternalQueryConnection(
                record.Offset,
                record.Type,
                sheetName,
                futureRecordType,
                dataSourceType,
                sourceTypeKind,
                ToSourceTypeName(dataSourceType, sourceTypeKind),
                connectionFlags,
                sourceSpecificFlags,
                queryOptions,
                editVersion,
                refreshedVersion,
                refreshableMinimumVersion,
                oleDbConnectionCount,
                futureByteCount,
                refreshIntervalMinutes,
                htmlFormat,
                parameterFlagCount,
                parameterFlagByteCount,
                parameterFlagByteCount == checked(parameterFlagCount * ParameterFlagByteLength));
            return true;
        }

        private static LegacyXlsExternalQueryConnectionSourceType ToSourceTypeKind(ushort sourceType) {
            switch (sourceType) {
                case 0x0001:
                    return LegacyXlsExternalQueryConnectionSourceType.Odbc;
                case 0x0002:
                    return LegacyXlsExternalQueryConnectionSourceType.Dao;
                case 0x0004:
                    return LegacyXlsExternalQueryConnectionSourceType.Web;
                case 0x0005:
                    return LegacyXlsExternalQueryConnectionSourceType.OleDb;
                case 0x0006:
                    return LegacyXlsExternalQueryConnectionSourceType.Text;
                case 0x0007:
                    return LegacyXlsExternalQueryConnectionSourceType.Ado;
                default:
                    return LegacyXlsExternalQueryConnectionSourceType.Unknown;
            }
        }

        private static string ToSourceTypeName(ushort sourceType, LegacyXlsExternalQueryConnectionSourceType kind) {
            return kind == LegacyXlsExternalQueryConnectionSourceType.Unknown
                ? $"DataSourceType:0x{sourceType:X4}"
                : kind.ToString();
        }

        private static void AddInvalidDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            BiffRecord record,
            string message) {
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                "XLS-BIFF-DBQUERYEXT-INVALID",
                "The DBQueryExt external query connection record could not be decoded. " + message,
                recordOffset: record.Offset,
                recordType: record.Type,
                detailCode: "DbQueryExtInvalid"));
        }
    }
}
