using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes worksheet print options, margins, page setup, and textual headers and footers.</summary>
    internal static class XlsbWorksheetPrintSettingsReader {
        private const string PrinterSettingsRelationshipSuffix = "/printerSettings";

        internal static XlsbPrintOptions ReadPrintOptions(XlsbRecord record) {
            if (record.Data.Length != 2) {
                throw new InvalidDataException($"The BrtPrintOptions record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            ushort flags = new XlsbBinaryCursor(record.Data).ReadUInt16();
            if ((flags & 0xFFE0) != 0) {
                throw new InvalidDataException($"The BrtPrintOptions record at offset {record.Offset} contains reserved flags.");
            }
            return new XlsbPrintOptions(
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0);
        }

        internal static XlsbPageMargins ReadMargins(XlsbRecord record) {
            if (record.Data.Length != 48) {
                throw new InvalidDataException($"The BrtMargins record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            return new XlsbPageMargins(
                ReadMargin(cursor, record, "left"),
                ReadMargin(cursor, record, "right"),
                ReadMargin(cursor, record, "top"),
                ReadMargin(cursor, record, "bottom"),
                ReadMargin(cursor, record, "header"),
                ReadMargin(cursor, record, "footer"));
        }

        internal static XlsbPageSetup ReadPageSetup(XlsbRecord record) {
            if (record.Data.Length < 38) {
                throw new InvalidDataException($"The BrtPageSetup record at offset {record.Offset} is truncated.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            uint paperSize = cursor.ReadUInt32();
            uint scale = cursor.ReadUInt32();
            uint horizontalDpi = cursor.ReadUInt32();
            uint verticalDpi = cursor.ReadUInt32();
            uint copies = cursor.ReadUInt32();
            int firstPageNumber = cursor.ReadInt32();
            uint fitToWidth = cursor.ReadUInt32();
            uint fitToHeight = cursor.ReadUInt32();
            ushort flags = cursor.ReadUInt16();
            if (paperSize >= int.MaxValue
                || (paperSize >= 119U && paperSize <= 256U)
                || (scale != 0U && (scale < 10U || scale > 400U))
                || copies > 32_767U
                || firstPageNumber < -32_765
                || firstPageNumber > 32_767
                || fitToWidth > 32_767U
                || fitToHeight > 32_767U
                || (flags & 0xF804) != 0) {
                throw new InvalidDataException($"The BrtPageSetup record at offset {record.Offset} contains invalid page settings or reserved flags.");
            }

            string? relationshipId = XlsbBinaryValueReader.ReadNullableWideString(cursor, 260);
            if (cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtPageSetup record at offset {record.Offset} has {cursor.Remaining} unexpected trailing payload bytes.");
            }

            return new XlsbPageSetup(
                paperSize,
                scale,
                horizontalDpi,
                verticalDpi,
                copies,
                firstPageNumber,
                fitToWidth,
                fitToHeight,
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0008) != 0,
                (flags & 0x0010) != 0,
                (flags & 0x0020) != 0,
                (flags & 0x0040) != 0,
                (flags & 0x0080) != 0,
                (flags & 0x0100) != 0,
                (XlsbPrintErrorMode)((flags >> 9) & 0x03),
                string.IsNullOrEmpty(relationshipId) ? null : relationshipId);
        }

        internal static void ValidatePrinterSettingsRelationship(
            XlsbRecord record,
            XlsbPageSetup pageSetup,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships) {
            string? relationshipId = pageSetup.PrinterSettingsRelationshipId;
            if (relationshipId == null) return;
            if (!relationships.TryGetValue(relationshipId, out XlsbPackageRelationship? relationship)
                || relationship.IsExternal
                || !relationship.Type.EndsWith(PrinterSettingsRelationshipSuffix, StringComparison.Ordinal)) {
                throw new InvalidDataException($"The BrtPageSetup record at offset {record.Offset} refers to missing or invalid printer-settings relationship '{relationshipId}'.");
            }
        }

        internal static XlsbHeaderFooter ReadHeaderFooter(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            ushort flags = cursor.ReadUInt16();
            if ((flags & 0xFFF0) != 0) {
                throw new InvalidDataException($"The BrtBeginHeaderFooter record at offset {record.Offset} contains reserved flags.");
            }
            string? oddHeader = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            string? oddFooter = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            string? evenHeader = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            string? evenFooter = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            string? firstHeader = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            string? firstFooter = XlsbBinaryValueReader.ReadNullableWideString(cursor, 255);
            if (cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtBeginHeaderFooter record at offset {record.Offset} has {cursor.Remaining} unexpected trailing payload bytes.");
            }
            return new XlsbHeaderFooter(
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0,
                oddHeader,
                oddFooter,
                evenHeader,
                evenFooter,
                firstHeader,
                firstFooter);
        }

        private static double ReadMargin(XlsbBinaryCursor cursor, XlsbRecord record, string detail) {
            double value = cursor.ReadDouble();
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value >= 49D) {
                throw new InvalidDataException($"The BrtMargins record at offset {record.Offset} contains invalid {detail} margin {value}.");
            }
            return value;
        }
    }
}
