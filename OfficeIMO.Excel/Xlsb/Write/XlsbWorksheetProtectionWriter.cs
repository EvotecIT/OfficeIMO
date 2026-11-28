using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes classic worksheet protection.</summary>
    internal static class XlsbWorksheetProtectionWriter {
        private const int BrtSheetProtection = 535;

        internal static void Write(Stream output, OfficeXmlNode protection) {
            ArgumentNullException.ThrowIfNull(output);
            if (protection.IsNone) return;
            XlsbRecordWriter.Write(output, BrtSheetProtection, CreatePayload(protection));
        }

        private static byte[] CreatePayload(OfficeXmlNode protection) {
            if (protection.FirstChild.Exists) {
                throw new NotSupportedException("Native XLSB generation does not support child content in worksheet protection.");
            }
            string? unsupported = XlsbWorksheetProtectionProjector.UnsupportedAttribute(protection);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet protection attribute '{unsupported}'.");
            }
            if (!XlsbWorksheetProtectionProjector.IsProtected(protection)) {
                throw new NotSupportedException("Native XLSB generation does not write inactive worksheet protection elements.");
            }
            if (!XlsbWorkbookProtectionProjector.TryParsePassword(protection.Attribute("password"), out ushort password)) {
                throw new NotSupportedException("Native XLSB generation requires a classic worksheet protection hash with at most four hexadecimal digits.");
            }

            using var output = new MemoryStream(66);
            XlsbPayloadWriter.WriteUInt16(output, password);
            XlsbPayloadWriter.WriteUInt32(output, 1U);
            foreach (bool allowed in XlsbWorksheetProtectionProjector.AllowedActions(protection)) {
                XlsbPayloadWriter.WriteUInt32(output, allowed ? 1U : 0U);
            }
            return output.ToArray();
        }
    }
}
