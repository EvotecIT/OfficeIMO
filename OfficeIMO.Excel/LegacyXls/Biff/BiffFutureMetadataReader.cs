using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffFutureMetadataReader {
        internal static bool TryCreateWorkbookRecord(
            BiffRecord record,
            out LegacyXlsWorkbookFutureMetadataRecord? futureRecord) {
            if (!TryGetKind(record.Type, out LegacyXlsWorkbookMetadataKind kind)) {
                futureRecord = null;
                return false;
            }

            (ushort? headerRecordType, ushort? headerFlags) = ReadHeader(record.Payload);
            futureRecord = new LegacyXlsWorkbookFutureMetadataRecord(
                kind,
                record.Offset,
                record.Type,
                record.Payload.Length,
                headerRecordType,
                headerFlags);
            return true;
        }

        internal static bool TryCreateWorksheetRecord(
            BiffRecord record,
            out LegacyXlsWorksheetFutureMetadataRecord? futureRecord) {
            if (!TryGetKind(record.Type, out LegacyXlsWorkbookMetadataKind kind)) {
                futureRecord = null;
                return false;
            }

            (ushort? headerRecordType, ushort? headerFlags) = ReadHeader(record.Payload);
            futureRecord = new LegacyXlsWorksheetFutureMetadataRecord(
                kind,
                record.Offset,
                record.Type,
                record.Payload.Length,
                headerRecordType,
                headerFlags);
            return true;
        }

        private static bool TryGetKind(ushort recordType, out LegacyXlsWorkbookMetadataKind kind) {
            switch ((BiffRecordType)recordType) {
                case BiffRecordType.RecalcId:
                    kind = LegacyXlsWorkbookMetadataKind.RecalculationIdentifier;
                    return true;
                case BiffRecordType.EntExU2:
                    kind = LegacyXlsWorkbookMetadataKind.ExtendedEncryption;
                    return true;
                case BiffRecordType.ContinueFrt:
                    kind = LegacyXlsWorkbookMetadataKind.FutureRecordContinuation;
                    return true;
                case BiffRecordType.Compat12:
                    kind = LegacyXlsWorkbookMetadataKind.Compatibility12;
                    return true;
                case BiffRecordType.NamePublish:
                    kind = LegacyXlsWorkbookMetadataKind.NamePublish;
                    return true;
                case BiffRecordType.NameCmt:
                    kind = LegacyXlsWorkbookMetadataKind.NameComment;
                    return true;
                case BiffRecordType.SortData:
                    kind = LegacyXlsWorkbookMetadataKind.SortData;
                    return true;
                case BiffRecordType.GuidTypeLib:
                    kind = LegacyXlsWorkbookMetadataKind.TypeLibraryGuid;
                    return true;
                case BiffRecordType.FnGrp12:
                    kind = LegacyXlsWorkbookMetadataKind.FunctionGroup12;
                    return true;
                case BiffRecordType.NameFnGrp12:
                    kind = LegacyXlsWorkbookMetadataKind.NameFunctionGroup12;
                    return true;
                case BiffRecordType.MtrSettings:
                    kind = LegacyXlsWorkbookMetadataKind.MultiThreadedRecalculationSettings;
                    return true;
                case BiffRecordType.CompressPictures:
                    kind = LegacyXlsWorkbookMetadataKind.CompressPictures;
                    return true;
                case BiffRecordType.HeaderFooter:
                    kind = LegacyXlsWorkbookMetadataKind.HeaderFooter;
                    return true;
                default:
                    kind = default;
                    return false;
            }
        }

        private static (ushort? HeaderRecordType, ushort? HeaderFlags) ReadHeader(byte[] payload) {
            ushort? headerRecordType = payload.Length >= 2
                ? BiffRecordReader.ReadUInt16(payload, 0)
                : null;
            ushort? headerFlags = payload.Length >= 4
                ? BiffRecordReader.ReadUInt16(payload, 2)
                : null;

            return (headerRecordType, headerFlags);
        }
    }
}
