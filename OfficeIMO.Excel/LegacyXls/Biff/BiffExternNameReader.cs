using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads ExternName records attached to a preceding SupBook supporting link.
    /// </summary>
    internal static class BiffExternNameReader {
        internal static bool TryRead(
            BiffRecord record,
            LegacyXlsExternalReferenceKind referenceKind,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsExternalName? externalName) {
            externalName = null;
            try {
                if (record.Payload.Length < 7) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-EXTERNNAME-SHORT",
                        "An ExternName record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return false;
                }

                ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 0);
                bool builtIn = (flags & 0x0001) != 0;
                bool wantsAdvise = (flags & 0x0002) != 0;
                bool wantsPicture = (flags & 0x0004) != 0;
                bool ole = (flags & 0x0008) != 0;
                bool oleLink = (flags & 0x0010) != 0;
                int cachedClipboardFormat = DecodeSignedTenBitValue((flags >> 5) & 0x03ff);
                bool icon = (flags & 0x8000) != 0;
                LegacyXlsExternalNameBodyKind bodyKind = GetBodyKind(referenceKind, ole, oleLink);
                int offset = 2;
                ushort oneBasedSheetIndex = 0;
                if (referenceKind == LegacyXlsExternalReferenceKind.AddIn) {
                    offset += 4; // skip AddinUdf.reserved
                } else {
                    oneBasedSheetIndex = BiffRecordReader.ReadUInt16(record.Payload, offset);
                    offset += 4; // skip ixals and reserved
                }

                string rawName = BiffStringReader.ReadShortUnicodeString(record.Payload, ref offset);
                string name = builtIn ? GetBuiltInName(rawName) ?? string.Empty : rawName;
                if (string.IsNullOrWhiteSpace(name)) {
                    return false;
                }

                int? localSheetIndex = oneBasedSheetIndex == 0 ? null : oneBasedSheetIndex - 1;
                externalName = new LegacyXlsExternalName(
                    name,
                    localSheetIndex,
                    builtIn,
                    wantsAdvise,
                    wantsPicture,
                    ole,
                    oleLink,
                    cachedClipboardFormat,
                    icon,
                    bodyKind);
                return true;
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-EXTERNNAME-INVALID",
                    $"An ExternName record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }
        }

        private static int DecodeSignedTenBitValue(int rawValue) {
            return rawValue >= 0x0200 ? rawValue - 0x0400 : rawValue;
        }

        private static LegacyXlsExternalNameBodyKind GetBodyKind(
            LegacyXlsExternalReferenceKind referenceKind,
            bool ole,
            bool oleLink) {
            if (referenceKind == LegacyXlsExternalReferenceKind.AddIn) {
                return LegacyXlsExternalNameBodyKind.AddInUdf;
            }

            if (referenceKind == LegacyXlsExternalReferenceKind.DdeOrOle) {
                if (ole) {
                    return LegacyXlsExternalNameBodyKind.DdeLinkNoOper;
                }

                if (oleLink) {
                    return LegacyXlsExternalNameBodyKind.OleDataItem;
                }

                return LegacyXlsExternalNameBodyKind.OleDdeLink;
            }

            return referenceKind == LegacyXlsExternalReferenceKind.ExternalWorkbook
                ? LegacyXlsExternalNameBodyKind.ExternalDefinedName
                : LegacyXlsExternalNameBodyKind.Unknown;
        }

        private static string? GetBuiltInName(string rawName) {
            if (rawName.Length != 1) {
                return null;
            }

            switch (rawName[0]) {
                case (char)0x06:
                    return "_xlnm.Print_Area";
                case (char)0x07:
                    return "_xlnm.Print_Titles";
                case (char)0x0d:
                    return "_FilterDatabase";
                default:
                    return null;
            }
        }
    }
}
