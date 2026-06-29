using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        internal static bool SupportsWorksheetPhoneticProperties(PhoneticProperties properties, out string? reason) {
            reason = null;
            if (properties.HasChildren) {
                reason = "worksheet phonetic settings";
                return false;
            }

            foreach (OpenXmlAttribute attribute in properties.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                    || (attribute.LocalName != "fontId"
                        && attribute.LocalName != "type"
                        && attribute.LocalName != "alignment")) {
                    reason = "worksheet phonetic settings";
                    return false;
                }
            }

            if (properties.HasAttributes && properties.FontId == null) {
                reason = "worksheet phonetic settings without fontId";
                return false;
            }

            if (properties.FontId?.Value > ushort.MaxValue) {
                reason = "worksheet phonetic font ids outside BIFF8 limits";
                return false;
            }

            if (!TryGetBiffPhoneticType(properties.Type?.Value, out _)) {
                reason = "worksheet phonetic type";
                return false;
            }

            if (!TryGetBiffPhoneticAlignment(properties.Alignment?.Value, out _)) {
                reason = "worksheet phonetic alignment";
                return false;
            }

            return true;
        }

        private static void WriteWorksheetPhoneticInfoRecords(Stream stream, ExcelSheet sheet) {
            if (TryCreateWorksheetPhoneticInfoPayload(sheet, out byte[]? payload)) {
                WriteRecord(stream, 0x00ed, payload!);
            }
        }

        private static bool TryCreateWorksheetPhoneticInfoPayload(ExcelSheet sheet, out byte[]? payload) {
            payload = null;
            PhoneticProperties? properties = sheet.WorksheetPart.Worksheet?
                .Elements<PhoneticProperties>()
                .FirstOrDefault(properties => properties.HasAttributes || properties.HasChildren);
            if (properties == null) {
                return false;
            }

            if (!SupportsWorksheetPhoneticProperties(properties, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason ?? "worksheet phonetic settings"}. Save as .xlsx or remove this feature before saving as .xls.");
            }

            if (properties.FontId == null) {
                return false;
            }

            ushort fontId = checked((ushort)properties.FontId.Value);
            TryGetBiffPhoneticType(properties.Type?.Value, out ushort typeCode);
            TryGetBiffPhoneticAlignment(properties.Alignment?.Value, out ushort alignmentCode);

            using var stream = new MemoryStream();
            WriteUInt16(stream, fontId);
            WriteUInt16(stream, (ushort)(typeCode | (alignmentCode << 2)));
            WriteUInt16(stream, 0);
            payload = stream.ToArray();
            return true;
        }

        private static bool TryGetBiffPhoneticType(PhoneticValues? value, out ushort code) {
            PhoneticValues resolved = value ?? PhoneticValues.FullWidthKatakana;
            if (resolved == PhoneticValues.HalfWidthKatakana) {
                code = 0;
                return true;
            }

            if (resolved == PhoneticValues.FullWidthKatakana) {
                code = 1;
                return true;
            }

            if (resolved == PhoneticValues.Hiragana) {
                code = 2;
                return true;
            }

            if (resolved == PhoneticValues.NoConversion) {
                code = 3;
                return true;
            }

            code = 0;
            return false;
        }

        private static bool TryGetBiffPhoneticAlignment(PhoneticAlignmentValues? value, out ushort code) {
            PhoneticAlignmentValues resolved = value ?? PhoneticAlignmentValues.Left;
            if (resolved == PhoneticAlignmentValues.NoControl) {
                code = 0;
                return true;
            }

            if (resolved == PhoneticAlignmentValues.Left) {
                code = 1;
                return true;
            }

            if (resolved == PhoneticAlignmentValues.Center) {
                code = 2;
                return true;
            }

            if (resolved == PhoneticAlignmentValues.Distributed) {
                code = 3;
                return true;
            }

            code = 0;
            return false;
        }
    }
}
