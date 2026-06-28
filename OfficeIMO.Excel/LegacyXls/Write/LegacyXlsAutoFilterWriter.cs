using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsAutoFilterWriter {
        private const byte ComparisonLessThan = 0x01;
        private const byte ComparisonEqual = 0x02;
        private const byte ComparisonLessThanOrEqual = 0x03;
        private const byte ComparisonGreaterThan = 0x04;
        private const byte ComparisonNotEqual = 0x05;
        private const byte ComparisonGreaterThanOrEqual = 0x06;

        internal static bool HasWorksheetAutoFilters(IReadOnlyList<ExcelSheet> sheets) {
            foreach (ExcelSheet sheet in sheets) {
                if (TryGetWorksheetAutoFilter(sheet, out _)) {
                    return true;
                }
            }

            return false;
        }

        internal static bool SupportsWorksheetAutoFilter(ExcelSheet sheet, out string? reason) {
            reason = null;
            if (!TryGetWorksheetAutoFilter(sheet, out AutoFilter? autoFilter)) {
                return true;
            }

            AutoFilter filter = autoFilter!;
            if (!TryGetAutoFilterRange(filter, out AutoFilterRange range, out reason)) {
                return false;
            }

            if (!SupportsAutoFilterMetadata(filter)) {
                reason = "AutoFilter metadata";
                return false;
            }

            if (filter.Elements<SortState>().Any()) {
                reason = "AutoFilter sort metadata";
                return false;
            }

            if (HasExtensionMetadata(filter)) {
                reason = "AutoFilter extension metadata";
                return false;
            }

            foreach (FilterColumn filterColumn in filter.Elements<FilterColumn>()) {
                if ((filterColumn.ColumnId?.Value ?? 0U) >= range.DropDownCount) {
                    reason = "AutoFilter column IDs outside the filter range";
                    return false;
                }

                if (!SupportsFilterColumn(filterColumn, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static byte[] CreateExternSheetPayload(IReadOnlyList<ExcelSheet> sheets) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)sheets.Count));
            for (int i = 0; i < sheets.Count; i++) {
                WriteUInt16(stream, 0);
                WriteUInt16(stream, checked((ushort)i));
                WriteUInt16(stream, checked((ushort)i));
            }

            return stream.ToArray();
        }

        internal static IReadOnlyList<byte[]> CreateDefinedNamePayloads(IReadOnlyList<ExcelSheet> sheets) {
            var payloads = new List<byte[]>();
            for (int i = 0; i < sheets.Count; i++) {
                if (!TryGetWorksheetAutoFilter(sheets[i], out AutoFilter? autoFilter)
                    || !TryGetAutoFilterRange(autoFilter!, out AutoFilterRange range, out _)) {
                    continue;
                }

                byte[] formula = BuildNameArea3dFormula(
                    checked((ushort)i),
                    checked((ushort)(range.FirstRow - 1)),
                    checked((ushort)(range.FirstColumn - 1)),
                    checked((ushort)(range.LastRow - 1)),
                    checked((ushort)(range.LastColumn - 1)));
                payloads.Add(BuildDefinedNamePayload(((char)0x0d).ToString(), formula, checked((ushort)(i + 1))));
            }

            return payloads;
        }

        internal static bool TryCreateAutoFilterInfoPayload(ExcelSheet sheet, out byte[]? payload) {
            payload = null;
            if (!TryGetWorksheetAutoFilter(sheet, out AutoFilter? autoFilter)
                || !TryGetAutoFilterRange(autoFilter!, out AutoFilterRange range, out _)) {
                return false;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, range.DropDownCount);
            payload = stream.ToArray();
            return true;
        }

        internal static IReadOnlyList<byte[]> CreateCriteriaPayloads(ExcelSheet sheet, ExcelDateSystem dateSystem) {
            if (!TryGetWorksheetAutoFilter(sheet, out AutoFilter? autoFilter)
                || !TryGetAutoFilterRange(autoFilter!, out AutoFilterRange range, out _)) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>();
            foreach (FilterColumn filterColumn in autoFilter!.Elements<FilterColumn>()) {
                if (TryCreateCriteriaPayload(filterColumn, range.DropDownCount, dateSystem, out byte[]? payload)) {
                    payloads.Add(payload!);
                }
            }

            return payloads;
        }

        private static bool TryGetWorksheetAutoFilter(ExcelSheet sheet, out AutoFilter? autoFilter) {
            autoFilter = sheet.WorksheetPart.Worksheet?.Elements<AutoFilter>().FirstOrDefault();
            return autoFilter != null;
        }

        private static bool TryGetAutoFilterRange(AutoFilter autoFilter, out AutoFilterRange range, out string? reason) {
            range = default;
            reason = null;
            string? reference = autoFilter.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "AutoFilter without a range";
                return false;
            }

            try {
                (int firstRow, int firstColumn, int lastRow, int lastColumn) = A1.ParseRange(reference!.Replace("$", string.Empty));
                if (firstRow < 1 || firstColumn < 1 || lastRow < firstRow || lastColumn < firstColumn) {
                    reason = "AutoFilter range";
                    return false;
                }

                if (lastRow > 65536 || lastColumn > 256) {
                    reason = "AutoFilter ranges outside BIFF8 worksheet limits";
                    return false;
                }

                int width = lastColumn - firstColumn + 1;
                if (width < 1 || width > 256) {
                    reason = "AutoFilter ranges wider than 256 columns";
                    return false;
                }

                range = new AutoFilterRange(firstRow, firstColumn, lastRow, lastColumn, checked((ushort)width));
                return true;
            } catch (Exception ex) when (ex is ArgumentException || ex is FormatException || ex is OverflowException) {
                reason = "AutoFilter range";
                return false;
            }
        }

        private static bool SupportsFilterColumn(FilterColumn filterColumn, out string? reason) {
            reason = null;
            if (!SupportsFilterColumnMetadata(filterColumn)) {
                reason = "AutoFilter column metadata";
                return false;
            }

            if (filterColumn.ColumnId?.Value == null || filterColumn.ColumnId.Value > 255U) {
                reason = "AutoFilter column IDs outside BIFF8 limits";
                return false;
            }

            if (filterColumn.HiddenButton?.Value == true || filterColumn.ShowButton?.Value == false) {
                reason = "AutoFilter dropdown control metadata";
                return false;
            }

            if (HasExtensionMetadata(filterColumn)) {
                reason = "AutoFilter extension metadata";
                return false;
            }

            int primaryChildCount = CountPrimaryFilterChildren(filterColumn);
            if (primaryChildCount == 0) {
                return true;
            }

            if (primaryChildCount > 1) {
                reason = "AutoFilter columns with multiple filter types";
                return false;
            }

            if (filterColumn.GetFirstChild<Filters>() is Filters filters) {
                return SupportsFilters(filters, out reason);
            }

            if (filterColumn.GetFirstChild<CustomFilters>() is CustomFilters customFilters) {
                return SupportsCustomFilters(customFilters, out reason);
            }

            if (filterColumn.GetFirstChild<Top10>() is Top10 top10) {
                return SupportsTop10(top10, out reason);
            }

            reason = "dynamic, color, icon, or extension AutoFilter criteria";
            return false;
        }

        private static bool HasExtensionMetadata(OpenXmlElement element) {
            return element.Elements<ExtensionList>().Any(extensionList => extensionList.Elements<Extension>().Any());
        }

        private static bool SupportsAutoFilterMetadata(AutoFilter autoFilter) {
            if (autoFilter.ChildElements.Any(child => child is not FilterColumn && child is not SortState && child is not ExtensionList)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in autoFilter.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "ref", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsFilterColumnMetadata(FilterColumn filterColumn) {
            if (filterColumn.ChildElements.Any(child =>
                    child is not Filters
                    && child is not CustomFilters
                    && child is not Top10
                    && child is not DynamicFilter
                    && child is not ColorFilter
                    && child is not IconFilter
                    && child is not ExtensionList)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in filterColumn.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "colId":
                    case "hiddenButton":
                    case "showButton":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsFilters(Filters filters, out string? reason) {
            reason = null;
            if (!SupportsFiltersMetadata(filters)) {
                reason = "AutoFilter criteria metadata";
                return false;
            }

            List<string> values = filters.Elements<Filter>()
                .Select(filter => filter.Val?.Value)
                .Where(value => value != null)
                .Select(value => value!)
                .ToList();
            List<DateGroupItem> dateGroups = filters.Elements<DateGroupItem>().ToList();
            if (dateGroups.Count > 0) {
                if (dateGroups.Count != 1
                    || filters.Blank?.Value == true
                    || values.Count > 0
                    || !LegacyXlsAutoFilterDateGroupRange.TryCreate(dateGroups[0], out _, out _)) {
                    reason = "AutoFilter date-group criteria";
                    return false;
                }

                return true;
            }

            if (filters.Blank?.Value == true && values.Count > 1) {
                reason = "AutoFilter blank criteria combined with more than one value";
                return false;
            }

            if (values.Count > 2) {
                reason = "AutoFilter equality lists with more than two values";
                return false;
            }

            foreach (string value in values) {
                if (!SupportsDoperString(value, out reason)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsFiltersMetadata(Filters filters) {
            if (filters.ChildElements.Any(child => child is not Filter && child is not DateGroupItem)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in filters.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "blank", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCustomFilters(CustomFilters customFilters, out string? reason) {
            reason = null;
            if (!SupportsCustomFiltersMetadata(customFilters)) {
                reason = "AutoFilter criteria metadata";
                return false;
            }

            List<CustomFilter> conditions = customFilters.Elements<CustomFilter>().ToList();
            if (conditions.Count == 0) {
                return true;
            }

            if (conditions.Count > 2) {
                reason = "AutoFilter custom filters with more than two conditions";
                return false;
            }

            foreach (CustomFilter condition in conditions) {
                if (!TryGetComparison(condition.Operator?.Value ?? FilterOperatorValues.Equal, out _)) {
                    reason = "AutoFilter custom filter operators outside BIFF8 limits";
                    return false;
                }

                string value = condition.Val?.Value ?? string.Empty;
                if (!IsProjectedNonBlankCondition(condition) && !SupportsDoperString(value, out reason)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCustomFiltersMetadata(CustomFilters customFilters) {
            if (customFilters.ChildElements.Any(child => child is not CustomFilter)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in customFilters.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "and", StringComparison.Ordinal)) {
                    return false;
                }
            }

            foreach (CustomFilter condition in customFilters.Elements<CustomFilter>()) {
                if (condition.HasChildren) {
                    return false;
                }

                foreach (OpenXmlAttribute attribute in condition.GetAttributes()) {
                    if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                        return false;
                    }

                    switch (attribute.LocalName) {
                        case "operator":
                        case "val":
                            break;
                        default:
                            return false;
                    }
                }
            }

            return true;
        }

        private static bool SupportsTop10(Top10 top10, out string? reason) {
            reason = null;
            if (!SupportsTop10Metadata(top10)) {
                reason = "AutoFilter criteria metadata";
                return false;
            }

            double value = top10.Val?.Value ?? 10d;
            double rounded = Math.Round(value);
            if (Math.Abs(value - rounded) > double.Epsilon || rounded < 1d || rounded > 500d) {
                reason = "AutoFilter top/bottom values outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool SupportsTop10Metadata(Top10 top10) {
            if (top10.HasChildren) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in top10.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "top":
                    case "percent":
                    case "val":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool TryCreateCriteriaPayload(FilterColumn filterColumn, ushort dropDownCount, ExcelDateSystem dateSystem, out byte[]? payload) {
            payload = null;
            uint columnId = filterColumn.ColumnId?.Value ?? 0U;
            if (columnId >= dropDownCount || columnId > 255U) {
                return false;
            }

            if (filterColumn.GetFirstChild<Filters>() is Filters filters) {
                payload = BuildFiltersPayload(checked((ushort)columnId), filters, dateSystem);
                return payload != null;
            }

            if (filterColumn.GetFirstChild<CustomFilters>() is CustomFilters customFilters) {
                payload = BuildCustomFiltersPayload(checked((ushort)columnId), customFilters);
                return payload != null;
            }

            if (filterColumn.GetFirstChild<Top10>() is Top10 top10) {
                payload = BuildTop10Payload(checked((ushort)columnId), top10);
                return payload != null;
            }

            return false;
        }

        private static byte[]? BuildFiltersPayload(ushort columnId, Filters filters, ExcelDateSystem dateSystem) {
            DateGroupItem? dateGroup = filters.Elements<DateGroupItem>().FirstOrDefault();
            if (dateGroup != null) {
                if (filters.Elements<DateGroupItem>().Skip(1).Any()
                    || filters.Blank?.Value == true
                    || filters.Elements<Filter>().Any()
                    || !LegacyXlsAutoFilterDateGroupRange.TryCreate(dateGroup, out DateTime start, out DateTime end)) {
                    return null;
                }

                using var dateGroupStream = new MemoryStream();
                WriteUInt16(dateGroupStream, columnId);
                WriteUInt16(dateGroupStream, 0x0000);
                WriteNumberDoper(dateGroupStream, ExcelDateSystemConverter.ToSerial(start, dateSystem), ComparisonGreaterThanOrEqual);
                WriteNumberDoper(dateGroupStream, ExcelDateSystemConverter.ToSerial(end, dateSystem), ComparisonLessThan);
                return dateGroupStream.ToArray();
            }

            List<string> values = filters.Elements<Filter>()
                .Select(filter => filter.Val?.Value)
                .Where(value => value != null)
                .Select(value => value!)
                .ToList();

            if (filters.Blank?.Value == true && values.Count == 0) {
                using var blankStream = new MemoryStream();
                WriteUInt16(blankStream, columnId);
                WriteUInt16(blankStream, 0x0004);
                WriteBlankDoper(blankStream, ComparisonEqual);
                WriteUnusedDoper(blankStream);
                return blankStream.ToArray();
            }

            if (filters.Blank?.Value == true && values.Count == 1) {
                using var blankOrValueStream = new MemoryStream();
                WriteUInt16(blankOrValueStream, columnId);
                WriteUInt16(blankOrValueStream, 0x0005);
                WriteStringDoper(blankOrValueStream, values[0], ComparisonEqual);
                WriteBlankDoper(blankOrValueStream, ComparisonEqual);
                WriteUnicodeStringNoCch(blankOrValueStream, values[0]);
                return blankOrValueStream.ToArray();
            }

            if (values.Count == 0 || values.Count > 2) {
                return null;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, columnId);
            WriteUInt16(stream, values.Count == 2 ? (ushort)0x0005 : (ushort)0x0004);
            WriteStringDoper(stream, values[0], ComparisonEqual);
            if (values.Count == 2) {
                WriteStringDoper(stream, values[1], ComparisonEqual);
            } else {
                WriteUnusedDoper(stream);
            }

            foreach (string value in values) {
                WriteUnicodeStringNoCch(stream, value);
            }

            return stream.ToArray();
        }

        private static byte[]? BuildCustomFiltersPayload(ushort columnId, CustomFilters customFilters) {
            List<CustomFilter> conditions = customFilters.Elements<CustomFilter>().ToList();
            if (conditions.Count == 0 || conditions.Count > 2) {
                return null;
            }

            if (conditions.Count == 1 && IsProjectedNonBlankCondition(conditions[0])) {
                using var nonBlankStream = new MemoryStream();
                WriteUInt16(nonBlankStream, columnId);
                WriteUInt16(nonBlankStream, 0x0004);
                WriteNonBlankDoper(nonBlankStream, ComparisonNotEqual);
                WriteUnusedDoper(nonBlankStream);
                return nonBlankStream.ToArray();
            }

            var tailStrings = new List<string>();
            bool hasString = false;
            using var stream = new MemoryStream();
            WriteUInt16(stream, columnId);
            ushort flags = hasString ? (ushort)0x0004 : (ushort)0;
            if (conditions.Count == 2 && customFilters.And?.Value != true) {
                flags |= 0x0001;
            }

            long flagsPosition = stream.Position;
            WriteUInt16(stream, flags);
            foreach (CustomFilter condition in conditions) {
                string value = condition.Val?.Value ?? string.Empty;
                byte comparison = TryGetComparison(condition.Operator?.Value ?? FilterOperatorValues.Equal, out byte mappedComparison)
                    ? mappedComparison
                    : ComparisonEqual;
                if (TryParseInvariantDouble(value, out double number)) {
                    WriteNumberDoper(stream, number, comparison);
                } else {
                    hasString = true;
                    WriteStringDoper(stream, value, comparison);
                    tailStrings.Add(value);
                }
            }

            if (conditions.Count == 1) {
                WriteUnusedDoper(stream);
            }

            long currentPosition = stream.Position;
            stream.Position = flagsPosition;
            WriteUInt16(stream, (ushort)(flags | (hasString ? 0x0004 : 0x0000)));
            stream.Position = currentPosition;

            foreach (string value in tailStrings) {
                WriteUnicodeStringNoCch(stream, value);
            }

            return stream.ToArray();
        }

        private static byte[]? BuildTop10Payload(ushort columnId, Top10 top10) {
            double value = top10.Val?.Value ?? 10d;
            ushort rounded = checked((ushort)Math.Round(value));
            if (rounded < 1 || rounded > 500) {
                return null;
            }

            ushort flags = (ushort)(0x0010 | ((rounded & 0x01ff) << 7));
            if (top10.Top?.Value != false) {
                flags |= 0x0020;
            }

            if (top10.Percent?.Value == true) {
                flags |= 0x0040;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, columnId);
            WriteUInt16(stream, flags);
            WriteUnusedDoper(stream);
            WriteUnusedDoper(stream);
            return stream.ToArray();
        }

        private static int CountPrimaryFilterChildren(FilterColumn filterColumn) {
            int count = 0;
            count += filterColumn.Elements<Filters>().Count();
            count += filterColumn.Elements<CustomFilters>().Count();
            count += filterColumn.Elements<Top10>().Count();
            count += filterColumn.Elements<DynamicFilter>().Count();
            count += filterColumn.Elements<ColorFilter>().Count();
            count += filterColumn.Elements<IconFilter>().Count();
            return count;
        }

        private static bool TryGetComparison(FilterOperatorValues value, out byte comparison) {
            if (value == FilterOperatorValues.LessThan) {
                comparison = ComparisonLessThan;
                return true;
            }

            if (value == FilterOperatorValues.Equal) {
                comparison = ComparisonEqual;
                return true;
            }

            if (value == FilterOperatorValues.LessThanOrEqual) {
                comparison = ComparisonLessThanOrEqual;
                return true;
            }

            if (value == FilterOperatorValues.GreaterThan) {
                comparison = ComparisonGreaterThan;
                return true;
            }

            if (value == FilterOperatorValues.NotEqual) {
                comparison = ComparisonNotEqual;
                return true;
            }

            if (value == FilterOperatorValues.GreaterThanOrEqual) {
                comparison = ComparisonGreaterThanOrEqual;
                return true;
            }

            comparison = 0;
            return false;
        }

        private static bool TryParseInvariantDouble(string value, out double number) {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        private static bool IsProjectedNonBlankCondition(CustomFilter condition) {
            return condition.Operator?.Value == FilterOperatorValues.NotEqual
                && string.Equals(condition.Val?.Value, " ", StringComparison.Ordinal);
        }

        private static bool SupportsDoperString(string value, out string? reason) {
            reason = null;
            if (value.Length > byte.MaxValue) {
                reason = "AutoFilter text values longer than 255 characters";
                return false;
            }

            return true;
        }

        private static byte[] BuildDefinedNamePayload(string name, byte[] formula, ushort localSheetIndex) {
            byte[] nameBytes = Encoding.ASCII.GetBytes(name);
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0021);
            stream.WriteByte(0);
            stream.WriteByte(checked((byte)name.Length));
            WriteUInt16(stream, checked((ushort)formula.Length));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, localSheetIndex);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.Write(nameBytes, 0, nameBytes.Length);
            stream.Write(formula, 0, formula.Length);
            return stream.ToArray();
        }

        private static byte[] BuildNameArea3dFormula(ushort externSheetIndex, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
            using var stream = new MemoryStream();
            stream.WriteByte(0x3b);
            WriteUInt16(stream, externSheetIndex);
            WriteUInt16(stream, firstRow);
            WriteUInt16(stream, lastRow);
            WriteUInt16(stream, firstColumn);
            WriteUInt16(stream, lastColumn);
            return stream.ToArray();
        }

        private static void WriteUnusedDoper(Stream stream) {
            stream.Write(new byte[10], 0, 10);
        }

        private static void WriteStringDoper(Stream stream, string value, byte comparison) {
            stream.WriteByte(0x06);
            stream.WriteByte(comparison);
            WriteUInt32(stream, 0);
            stream.WriteByte(checked((byte)value.Length));
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
        }

        private static void WriteNumberDoper(Stream stream, double value, byte comparison) {
            stream.WriteByte(0x04);
            stream.WriteByte(comparison);
            byte[] valueBytes = BitConverter.GetBytes(value);
            stream.Write(valueBytes, 0, valueBytes.Length);
        }

        private static void WriteBlankDoper(Stream stream, byte comparison) {
            stream.WriteByte(0x0c);
            stream.WriteByte(comparison);
            stream.Write(new byte[8], 0, 8);
        }

        private static void WriteNonBlankDoper(Stream stream, byte comparison) {
            stream.WriteByte(0x0e);
            stream.WriteByte(comparison);
            stream.Write(new byte[8], 0, 8);
        }

        private static void WriteUnicodeStringNoCch(Stream stream, string value) {
            byte[] bytes;
            if (CanUseCompressedString(value)) {
                stream.WriteByte(0);
                bytes = Encoding.ASCII.GetBytes(value);
            } else {
                stream.WriteByte(1);
                bytes = Encoding.Unicode.GetBytes(value);
            }

            stream.Write(bytes, 0, bytes.Length);
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
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        private readonly struct AutoFilterRange {
            internal AutoFilterRange(int firstRow, int firstColumn, int lastRow, int lastColumn, ushort dropDownCount) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
                DropDownCount = dropDownCount;
            }

            internal int FirstRow { get; }
            internal int FirstColumn { get; }
            internal int LastRow { get; }
            internal int LastColumn { get; }
            internal ushort DropDownCount { get; }
        }
    }
}
