using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes classic workbook protection as BrtBookProtection.</summary>
    internal static class XlsbWorkbookProtectionWriter {
        private const int BrtBookProtection = 534;

        internal static void Write(Stream output, WorkbookProtection? protection) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (protection == null) return;
            XlsbRecordWriter.Write(output, BrtBookProtection, CreatePayload(protection));
        }

        internal static void Validate(WorkbookProtection? protection) {
            if (protection == null) return;
            CreatePayload(protection);
        }

        private static byte[] CreatePayload(WorkbookProtection protection) {
            if (protection.HasChildren) {
                throw new NotSupportedException("Native XLSB generation does not support child content in workbook protection.");
            }
            EnsureOnlyAttributes(protection,
                "workbookPassword", "revisionsPassword", "lockStructure", "lockWindows", "lockRevision");
            if (!XlsbWorkbookProtectionProjector.TryParsePassword(protection.WorkbookPassword?.Value, out ushort workbookPassword)
                || !XlsbWorkbookProtectionProjector.TryParsePassword(protection.RevisionsPassword?.Value, out ushort revisionsPassword)) {
                throw new NotSupportedException("Native XLSB generation requires classic workbook protection hashes with at most four hexadecimal digits.");
            }

            ushort flags = (ushort)((protection.LockStructure?.Value == true ? 0x0001 : 0)
                | (protection.LockWindows?.Value == true ? 0x0002 : 0)
                | (protection.LockRevision?.Value == true ? 0x0004 : 0));
            return new[] {
                (byte)workbookPassword,
                (byte)(workbookPassword >> 8),
                (byte)revisionsPassword,
                (byte)(revisionsPassword >> 8),
                (byte)flags,
                (byte)(flags >> 8)
            };
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook protection attribute '{unsupported.Value.LocalName}'.");
            }
        }
    }
}
