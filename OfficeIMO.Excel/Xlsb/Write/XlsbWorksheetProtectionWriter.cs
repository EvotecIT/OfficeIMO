using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes classic worksheet protection.</summary>
    internal static class XlsbWorksheetProtectionWriter {
        private const int BrtSheetProtection = 535;

        internal static void Write(Stream output, SheetProtection? protection) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (protection == null) return;
            XlsbRecordWriter.Write(output, BrtSheetProtection, CreatePayload(protection));
        }

        internal static void Validate(SheetProtection? protection) {
            if (protection != null) CreatePayload(protection);
        }

        private static byte[] CreatePayload(SheetProtection protection) {
            if (protection.HasChildren) {
                throw new NotSupportedException("Native XLSB generation does not support child content in worksheet protection.");
            }
            EnsureOnlyAttributes(protection,
                "password", "sheet", "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
                "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows",
                "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells");
            if (protection.Sheet?.Value == false) {
                throw new NotSupportedException("Native XLSB generation does not write inactive worksheet protection elements.");
            }
            if (!XlsbWorksheetProtectionProjector.TryParsePassword(protection.Password?.Value, out ushort password)) {
                throw new NotSupportedException("Native XLSB generation requires a classic worksheet protection hash with at most four hexadecimal digits.");
            }

            using var output = new MemoryStream(66);
            WriteUInt16(output, password);
            WriteBoolean(output, true);
            WriteBoolean(output, IsAllowed(protection.Objects, lockedWhenOmitted: false));
            WriteBoolean(output, IsAllowed(protection.Scenarios, lockedWhenOmitted: false));
            WriteBoolean(output, IsAllowed(protection.FormatCells, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.FormatColumns, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.FormatRows, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.InsertColumns, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.InsertRows, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.InsertHyperlinks, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.DeleteColumns, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.DeleteRows, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.SelectLockedCells, lockedWhenOmitted: false));
            WriteBoolean(output, IsAllowed(protection.Sort, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.AutoFilter, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.PivotTables, lockedWhenOmitted: true));
            WriteBoolean(output, IsAllowed(protection.SelectUnlockedCells, lockedWhenOmitted: false));
            return output.ToArray();
        }

        private static bool IsAllowed(BooleanValue? protectionFlag, bool lockedWhenOmitted) =>
            !(protectionFlag?.Value ?? lockedWhenOmitted);

        private static void EnsureOnlyAttributes(OpenXmlElement element, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet protection attribute '{unsupported.Value.LocalName}'.");
            }
        }

        private static void WriteBoolean(Stream output, bool value) => WriteUInt32(output, value ? 1U : 0U);

        private static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }
    }
}
