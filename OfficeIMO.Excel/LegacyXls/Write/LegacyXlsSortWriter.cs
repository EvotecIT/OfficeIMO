using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsSortWriter {
        internal static bool SupportsWorksheetSortState(ExcelSheet sheet, out string? reason) {
            reason = null;
            SortState? sortState = sheet.WorksheetPart.Worksheet?.GetFirstChild<SortState>();
            if (sortState == null) {
                return true;
            }

            if (!SupportsSortStateMetadata(sortState)) {
                reason = "sort states with unsupported metadata";
                return false;
            }

            SortCondition[] conditions = sortState.Elements<SortCondition>().ToArray();
            if (conditions.Length > 3) {
                reason = "sort states with more than three sort keys";
                return false;
            }

            for (int i = 0; i < conditions.Length; i++) {
                SortCondition condition = conditions[i];
                if (!SupportsSortCondition(condition, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static byte[]? CreateSortPayload(ExcelSheet sheet) {
            SortState? sortState = sheet.WorksheetPart.Worksheet?.GetFirstChild<SortState>();
            if (sortState == null) {
                return null;
            }

            SortCondition[] conditions = sortState.Elements<SortCondition>().ToArray();
            if (conditions.Length == 0) {
                return null;
            }

            ushort flags = 0;
            if (sortState.ColumnSort?.Value == true) {
                flags |= 0x0001;
            }

            if (conditions.Length > 0 && conditions[0].Descending?.Value == true) {
                flags |= 0x0002;
            }

            if (conditions.Length > 1 && conditions[1].Descending?.Value == true) {
                flags |= 0x0004;
            }

            if (conditions.Length > 2 && conditions[2].Descending?.Value == true) {
                flags |= 0x0008;
            }

            if (sortState.CaseSensitive?.Value == true) {
                flags |= 0x0010;
            }

            if (sortState.SortMethod?.Value == SortMethodValues.PinYin) {
                flags |= 0x0400;
            }

            string key1 = GetSortKey(conditions, 0);
            string key2 = GetSortKey(conditions, 1);
            string key3 = GetSortKey(conditions, 2);

            using var stream = new MemoryStream();
            WriteUInt16(stream, flags);
            stream.WriteByte(checked((byte)key1.Length));
            stream.WriteByte(checked((byte)key2.Length));
            stream.WriteByte(checked((byte)key3.Length));
            WriteUnicodeStringNoCch(stream, key1);
            WriteUnicodeStringNoCch(stream, key2);
            WriteUnicodeStringNoCch(stream, key3);
            stream.WriteByte(0);
            return stream.ToArray();
        }

        private static bool SupportsSortCondition(SortCondition condition, out string? reason) {
            reason = null;
            if (!SupportsSortConditionMetadata(condition)) {
                reason = "sort states with unsupported metadata";
                return false;
            }

            if (condition.SortBy?.Value != null && condition.SortBy.Value != SortByValues.Value) {
                reason = "sort states with color or icon sort conditions";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(condition.CustomList?.Value)) {
                reason = "sort states with custom-list sort conditions";
                return false;
            }

            if (condition.FormatId != null || condition.IconSet != null || condition.IconId != null) {
                reason = "sort states with color or icon sort conditions";
                return false;
            }

            string? reference = condition.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "sort states with missing sort-key references";
                return false;
            }

            if (reference!.Length > byte.MaxValue) {
                reason = "sort states with sort-key references longer than 255 characters";
                return false;
            }

            return true;
        }

        private static bool SupportsSortStateMetadata(SortState sortState) {
            if (sortState.ChildElements.Any(child => child is not SortCondition)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in sortState.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "ref":
                    case "columnSort":
                    case "caseSensitive":
                    case "sortMethod":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsSortConditionMetadata(SortCondition condition) {
            if (condition.HasChildren) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in condition.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "descending":
                    case "sortBy":
                    case "ref":
                    case "customList":
                    case "dxfId":
                    case "iconSet":
                    case "iconId":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static string GetSortKey(IReadOnlyList<SortCondition> conditions, int index) {
            return index < conditions.Count
                ? conditions[index].Reference?.Value ?? string.Empty
                : string.Empty;
        }

        private static void WriteUnicodeStringNoCch(Stream stream, string value) {
            byte[] encoded = EncodeUnicodeString(value, out byte flags);
            stream.WriteByte(flags);
            stream.Write(encoded, 0, encoded.Length);
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)(value >> 8));
        }
    }
}
