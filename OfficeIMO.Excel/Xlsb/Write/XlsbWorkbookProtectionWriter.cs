using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes classic workbook protection as BrtBookProtection.</summary>
    internal static class XlsbWorkbookProtectionWriter {
        private const int BrtBookProtection = 534;

        internal static void Write(Stream output, OfficeXmlNode protection) {
            ArgumentNullException.ThrowIfNull(output);
            if (protection.IsNone) return;
            XlsbRecordWriter.Write(output, BrtBookProtection, CreatePayload(protection));
        }

        internal static void Validate(OfficeXmlNode protection) {
            if (protection.IsNone) return;
            CreatePayload(protection);
        }

        private static byte[] CreatePayload(OfficeXmlNode protection) {
            if (protection.FirstChild.Exists) {
                throw new NotSupportedException("Native XLSB generation does not support child content in workbook protection.");
            }
            string? unsupported = XlsbWorkbookProtectionProjector.UnsupportedAttribute(protection);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook protection attribute '{unsupported}'.");
            }
            if (!XlsbWorkbookProtectionProjector.TryParsePassword(protection.Attribute("workbookPassword"), out ushort workbookPassword)
                || !XlsbWorkbookProtectionProjector.TryParsePassword(protection.Attribute("revisionsPassword"), out ushort revisionsPassword)) {
                throw new NotSupportedException("Native XLSB generation requires classic workbook protection hashes with at most four hexadecimal digits.");
            }

            ushort flags = (ushort)((XlsbWorkbookProtectionProjector.ReadLock(protection, "lockStructure") ? 0x0001 : 0)
                | (XlsbWorkbookProtectionProjector.ReadLock(protection, "lockWindows") ? 0x0002 : 0)
                | (XlsbWorkbookProtectionProjector.ReadLock(protection, "lockRevision") ? 0x0004 : 0));
            return new[] {
                (byte)workbookPassword,
                (byte)(workbookPassword >> 8),
                (byte)revisionsPassword,
                (byte)(revisionsPassword >> 8),
                (byte)flags,
                (byte)(flags >> 8)
            };
        }
    }
}
