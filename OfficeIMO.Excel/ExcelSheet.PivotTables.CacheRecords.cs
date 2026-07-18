using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private List<PivotFieldValues> BuildPivotFieldValueMap(int fieldCount, int firstDataRow, int lastDataRow, int firstColumn,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap, IReadOnlyList<bool>? collectFieldValues = null) {
            var fieldValues = new List<PivotFieldValue>[fieldCount];
            var seenValues = new HashSet<string>[fieldCount];
            var fieldGroupings = new ExcelPivotGrouping?[fieldCount];
            var fieldsToCollect = new List<int>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                fieldValues[field] = new List<PivotFieldValue>();
                groupingMap.TryGetValue(field, out fieldGroupings[field]);
                bool collectValues = ShouldCollectPivotSharedItems(field, collectFieldValues);
                if (collectValues) {
                    seenValues[field] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    fieldsToCollect.Add(field);
                }
            }

            if (fieldsToCollect.Count > 0) {
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    for (int i = 0; i < fieldsToCollect.Count; i++) {
                        int field = fieldsToCollect[i];
                        int column = firstColumn + field;
                        var grouping = fieldGroupings[field];
                        var value = GetPivotFieldValue(row, column, grouping);
                        string text = value.Text;
                        if (seenValues[field]!.Add(text)) {
                            fieldValues[field].Add(value);
                        }
                    }
                }
            }

            var maps = new List<PivotFieldValues>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                maps.Add(new PivotFieldValues(fieldValues[field]));
            }

            return maps;
        }

        private List<PivotFieldValues> BuildPivotFieldValueMap(IExcelSheetTabularRowSource source, int fieldCount, int firstDataRow, int lastDataRow, int firstColumn,
            IReadOnlyList<bool>? collectFieldValues = null) {
            var fieldValues = new List<PivotFieldValue>[fieldCount];
            var seenValues = new HashSet<string>[fieldCount];
            var fieldsToCollect = new List<int>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                fieldValues[field] = new List<PivotFieldValue>();
                if (ShouldCollectPivotSharedItems(field, collectFieldValues)) {
                    seenValues[field] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    fieldsToCollect.Add(field);
                }
            }

            int firstSourceRow = firstDataRow - 2;
            int lastSourceRow = lastDataRow - 2;
            int sourceColumnOffset = firstColumn - 1;
            object?[]? flatValues = null;
            int flatColumnCount = 0;
            if (source.TryGetFlatValues(out var sourceFlatValues, out int sourceFlatColumnCount)
                && sourceFlatColumnCount >= sourceColumnOffset + fieldCount) {
                flatValues = sourceFlatValues;
                flatColumnCount = sourceFlatColumnCount;
            }

            if (fieldsToCollect.Count > 0) {
                for (int row = firstSourceRow; row <= lastSourceRow; row++) {
                    object?[]? rowValues = null;
                    bool hasBufferedRow = flatValues == null
                        && source.TryGetBufferedRow(row, out rowValues)
                        && rowValues != null
                        && rowValues.Length >= sourceColumnOffset + fieldCount;
                    for (int i = 0; i < fieldsToCollect.Count; i++) {
                        int field = fieldsToCollect[i];
                        int sourceColumnIndex = sourceColumnOffset + field;
                        object? rawValue = flatValues != null
                            ? flatValues[row * flatColumnCount + sourceColumnIndex]
                            : hasBufferedRow
                                ? rowValues![sourceColumnIndex]
                                : source.GetValue(row, sourceColumnIndex);
                        var value = GetPivotFieldValue(NormalizePivotRowSourceValue(rawValue));
                        if (seenValues[field]!.Add(value.Text)) {
                            fieldValues[field].Add(value);
                        }
                    }
                }
            }

            var maps = new List<PivotFieldValues>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                maps.Add(new PivotFieldValues(fieldValues[field]));
            }

            return maps;
        }

        private static bool ShouldCollectPivotSharedItems(int fieldIndex, IReadOnlyList<bool>? collectFieldValues)
            => collectFieldValues == null || fieldIndex < 0 || fieldIndex >= collectFieldValues.Count || collectFieldValues[fieldIndex];

        private static bool ShouldEmbedPivotCacheRecords(
            int sourceRecordCount,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<ExcelPivotCalculatedField> calculatedFields,
            IReadOnlyList<ExcelPivotFilter> pivotFilters,
            IReadOnlyDictionary<int, ExcelPivotFieldOptions>? fieldOptionMap)
        {
            if (sourceRecordCount <= 0 || sourceRecordCount > EmbeddedPivotCacheRecordRowLimit) {
                return false;
            }

            // Simple source-range pivots can refresh from the worksheet. Embed cache
            // records only for bounded scenarios that introduce cache-only fields.
            return groupingMap.Count != 0
                || generatedFields.Count != 0
                || calculatedFields.Count != 0;
        }

        private List<PivotFieldValues> BuildGeneratedPivotFieldValueMap(
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            int firstDataRow,
            int lastDataRow,
            int firstColumn) {
            var maps = new List<PivotFieldValues>(generatedFields.Count);
            foreach (var generatedField in generatedFields) {
                var values = new List<PivotFieldValue>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int column = firstColumn + generatedField.SourceIndex;
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    var value = GetGeneratedPivotDateFieldValue(row, column, generatedField.GroupBy);
                    if (seen.Add(value.Text)) {
                        values.Add(value);
                    }
                }

                maps.Add(new PivotFieldValues(values));
            }

            return maps;
        }

        private PivotCacheRecords BuildPivotCacheRecords(
            int fieldCount,
            int firstDataRow,
            int lastDataRow,
            int firstColumn,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap,
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<bool> sharedItemFields,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<PivotFieldValues> generatedFieldValueMap,
            int calculatedFieldCount) {
            int recordCount = Math.Max(0, lastDataRow - firstDataRow + 1);
            var records = new PivotCacheRecords { Count = (uint)recordCount };
            var lookups = BuildPivotRecordSharedItemLookups(fieldValueMap, sharedItemFields);
            var generatedLookups = BuildPivotRecordSharedItemLookups(generatedFieldValueMap, null);

            for (int row = firstDataRow; row <= lastDataRow; row++) {
                var record = new PivotCacheRecord();
                for (int field = 0; field < fieldCount; field++) {
                    groupingMap.TryGetValue(field, out var grouping);
                    record.Append(CreatePivotCacheRecordItem(GetPivotFieldValue(row, firstColumn + field, grouping), lookups[field]));
                }

                AppendGeneratedPivotCacheRecordItems(record, generatedFields, generatedLookups, row, firstColumn);
                AppendCalculatedPivotCacheRecordItems(record, calculatedFieldCount);
                records.Append(record);
            }

            return records;
        }

        private PivotCacheRecords BuildPivotCacheRecords(
            IExcelSheetTabularRowSource source,
            int fieldCount,
            int firstDataRow,
            int lastDataRow,
            int firstColumn,
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<bool> sharedItemFields,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<PivotFieldValues> generatedFieldValueMap,
            int calculatedFieldCount) {
            int recordCount = Math.Max(0, lastDataRow - firstDataRow + 1);
            var records = new PivotCacheRecords { Count = (uint)recordCount };
            var lookups = BuildPivotRecordSharedItemLookups(fieldValueMap, sharedItemFields);
            var generatedLookups = BuildPivotRecordSharedItemLookups(generatedFieldValueMap, null);
            int firstSourceRow = firstDataRow - 2;
            int lastSourceRow = lastDataRow - 2;
            int sourceColumnOffset = firstColumn - 1;
            object?[]? flatValues = null;
            int flatColumnCount = 0;
            if (source.TryGetFlatValues(out var sourceFlatValues, out int sourceFlatColumnCount)
                && sourceFlatColumnCount >= sourceColumnOffset + fieldCount) {
                flatValues = sourceFlatValues;
                flatColumnCount = sourceFlatColumnCount;
            }

            for (int row = firstSourceRow; row <= lastSourceRow; row++) {
                var record = new PivotCacheRecord();
                object?[]? rowValues = null;
                bool hasBufferedRow = flatValues == null
                    && source.TryGetBufferedRow(row, out rowValues)
                    && rowValues != null
                    && rowValues.Length >= sourceColumnOffset + fieldCount;
                for (int field = 0; field < fieldCount; field++) {
                    int sourceColumnIndex = sourceColumnOffset + field;
                    object? rawValue = flatValues != null
                        ? flatValues[row * flatColumnCount + sourceColumnIndex]
                        : hasBufferedRow
                            ? rowValues![sourceColumnIndex]
                            : source.GetValue(row, sourceColumnIndex);
                    record.Append(CreatePivotCacheRecordItem(
                        GetPivotFieldValue(NormalizePivotRowSourceValue(rawValue)),
                        lookups[field]));
                }

                AppendGeneratedPivotCacheRecordItems(record, generatedFields, generatedLookups, row + 2, firstColumn);
                AppendCalculatedPivotCacheRecordItems(record, calculatedFieldCount);
                records.Append(record);
            }

            return records;
        }

        private void WritePivotCacheRecords(
            PivotTableCacheRecordsPart cacheRecordsPart,
            IExcelSheetTabularRowSource source,
            int fieldCount,
            int firstDataRow,
            int lastDataRow,
            int firstColumn,
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<bool> sharedItemFields,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<PivotFieldValues> generatedFieldValueMap,
            int calculatedFieldCount) {
            int recordCount = Math.Max(0, lastDataRow - firstDataRow + 1);
            var lookups = BuildPivotRecordSharedItemLookups(fieldValueMap, sharedItemFields);
            var generatedLookups = BuildPivotRecordSharedItemLookups(generatedFieldValueMap, null);
            int firstSourceRow = firstDataRow - 2;
            int lastSourceRow = lastDataRow - 2;
            int sourceColumnOffset = firstColumn - 1;
            object?[]? flatValues = null;
            int flatColumnCount = 0;
            if (source.TryGetFlatValues(out var sourceFlatValues, out int sourceFlatColumnCount)
                && sourceFlatColumnCount >= sourceColumnOffset + fieldCount) {
                flatValues = sourceFlatValues;
                flatColumnCount = sourceFlatColumnCount;
            }

            using (var stream = cacheRecordsPart.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 65536)) {
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
                writer.Write(recordCount.ToString(CultureInfo.InvariantCulture));
                writer.Write("\">");
                for (int row = firstSourceRow; row <= lastSourceRow; row++) {
                    writer.Write("<r>");
                    object?[]? rowValues = null;
                    bool hasBufferedRow = flatValues == null
                        && source.TryGetBufferedRow(row, out rowValues)
                        && rowValues != null
                        && rowValues.Length >= sourceColumnOffset + fieldCount;
                    for (int field = 0; field < fieldCount; field++) {
                        int sourceColumnIndex = sourceColumnOffset + field;
                        object? rawValue = flatValues != null
                            ? flatValues[row * flatColumnCount + sourceColumnIndex]
                            : hasBufferedRow
                                ? rowValues![sourceColumnIndex]
                                : source.GetValue(row, sourceColumnIndex);
                        WritePivotCacheRecordItemXml(
                            writer,
                            GetPivotFieldValue(NormalizePivotRowSourceValue(rawValue)),
                            lookups[field]);
                    }

                    WriteGeneratedPivotCacheRecordItemsXml(writer, generatedFields, generatedLookups, row + 2, firstColumn);
                    WriteCalculatedPivotCacheRecordItemsXml(writer, calculatedFieldCount);
                    writer.Write("</r>");
                }

                writer.Write("</pivotCacheRecords>");
            }

            ExcelDocument.MarkPivotCacheRecordsPartAsRawWritten(cacheRecordsPart);
        }

        private static object? NormalizePivotRowSourceValue(object? value)
            => value == DBNull.Value ? null : value;

        private static Dictionary<string, uint>?[] BuildPivotRecordSharedItemLookups(
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<bool>? sharedItemFields) {
            var lookups = new Dictionary<string, uint>?[fieldValueMap.Count];
            for (int field = 0; field < fieldValueMap.Count; field++) {
                if (sharedItemFields != null
                    && (field < 0 || field >= sharedItemFields.Count || !sharedItemFields[field])) {
                    continue;
                }

                var items = fieldValueMap[field].Items;
                if (items.Count == 0) {
                    continue;
                }

                var lookup = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
                for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
                    lookup[items[itemIndex].Text] = (uint)itemIndex;
                }

                lookups[field] = lookup;
            }

            return lookups;
        }

        private void AppendGeneratedPivotCacheRecordItems(
            PivotCacheRecord record,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<Dictionary<string, uint>?> generatedLookups,
            int row,
            int firstColumn) {
            for (int i = 0; i < generatedFields.Count; i++) {
                var generatedField = generatedFields[i];
                var value = GetGeneratedPivotDateFieldValue(row, firstColumn + generatedField.SourceIndex, generatedField.GroupBy);
                record.Append(CreatePivotCacheRecordItem(value, generatedLookups[i]));
            }
        }

        private static void AppendCalculatedPivotCacheRecordItems(PivotCacheRecord record, int calculatedFieldCount) {
            for (int i = 0; i < calculatedFieldCount; i++) {
                record.Append(new MissingItem());
            }
        }

        private void WriteGeneratedPivotCacheRecordItems(
            OpenXmlWriter writer,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<Dictionary<string, uint>?> generatedLookups,
            int row,
            int firstColumn) {
            for (int i = 0; i < generatedFields.Count; i++) {
                var generatedField = generatedFields[i];
                var value = GetGeneratedPivotDateFieldValue(row, firstColumn + generatedField.SourceIndex, generatedField.GroupBy);
                writer.WriteElement(CreatePivotCacheRecordItem(value, generatedLookups[i]));
            }
        }

        private static void WriteCalculatedPivotCacheRecordItems(OpenXmlWriter writer, int calculatedFieldCount) {
            for (int i = 0; i < calculatedFieldCount; i++) {
                writer.WriteElement(new MissingItem());
            }
        }

        private void WriteGeneratedPivotCacheRecordItemsXml(
            TextWriter writer,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<Dictionary<string, uint>?> generatedLookups,
            int row,
            int firstColumn) {
            for (int i = 0; i < generatedFields.Count; i++) {
                var generatedField = generatedFields[i];
                var value = GetGeneratedPivotDateFieldValue(row, firstColumn + generatedField.SourceIndex, generatedField.GroupBy);
                WritePivotCacheRecordItemXml(writer, value, generatedLookups[i]);
            }
        }

        private static void WriteCalculatedPivotCacheRecordItemsXml(TextWriter writer, int calculatedFieldCount) {
            for (int i = 0; i < calculatedFieldCount; i++) {
                writer.Write("<m/>");
            }
        }

        private static void WritePivotCacheRecordItemXml(TextWriter writer, PivotFieldValue value, IReadOnlyDictionary<string, uint>? sharedItems) {
            if (sharedItems != null && sharedItems.TryGetValue(value.Text, out uint index)) {
                writer.Write("<x v=\"");
                writer.Write(index.ToString(CultureInfo.InvariantCulture));
                writer.Write("\"/>");
                return;
            }

            switch (value.Kind) {
                case PivotFieldValueKind.Blank:
                    writer.Write("<m/>");
                    break;
                case PivotFieldValueKind.Boolean:
                    writer.Write(value.Boolean!.Value ? "<b v=\"1\"/>" : "<b v=\"0\"/>");
                    break;
                case PivotFieldValueKind.Number:
                    writer.Write("<n v=\"");
                    writer.Write(value.Text);
                    writer.Write("\"/>");
                    break;
                case PivotFieldValueKind.Date:
                    writer.Write("<d v=\"");
                    WriteXmlAttributeEscaped(writer, value.Text);
                    writer.Write("\"/>");
                    break;
                default:
                    writer.Write("<s v=\"");
                    WriteXmlAttributeEscaped(writer, value.Text);
                    writer.Write("\"/>");
                    break;
            }
        }

        private static void WriteXmlAttributeEscaped(TextWriter writer, string value) {
            for (int i = 0; i < value.Length; i++) {
                char current = value[i];
                switch (current) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '"':
                        writer.Write("&quot;");
                        break;
                    case '\r':
                        writer.Write("&#xD;");
                        break;
                    case '\n':
                        writer.Write("&#xA;");
                        break;
                    case '\t':
                        writer.Write("&#x9;");
                        break;
                    default:
                        if (char.IsHighSurrogate(current)) {
                            if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) {
                                writer.Write(current);
                                writer.Write(value[++i]);
                            }

                            break;
                        }

                        if (char.IsLowSurrogate(current)) {
                            break;
                        }

                        if (IsLegalXmlChar(current)) {
                            writer.Write(current);
                        }
                        break;
                }
            }
        }

        private static bool IsLegalXmlChar(char value)
            => value == 0x9
               || value == 0xA
               || value == 0xD
               || (value >= 0x20 && value <= 0xD7FF)
               || (value >= 0xE000 && value <= 0xFFFD);

        private static OpenXmlElement CreatePivotCacheRecordItem(PivotFieldValue value, IReadOnlyDictionary<string, uint>? sharedItems) {
            if (sharedItems != null && sharedItems.TryGetValue(value.Text, out uint index)) {
                return new FieldItem { Val = index };
            }

            return value.Kind switch {
                PivotFieldValueKind.Blank => new MissingItem(),
                PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                _ => new StringItem { Val = value.Text }
            };
        }

        private PivotFieldValue GetPivotFieldValue(int row, int column, ExcelPivotGrouping? grouping) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            if (grouping?.IsDateGrouping == true) {
                if (TryGetPivotDateValue(row, column, text, out var date)) {
                    return PivotFieldValue.FromDate(date);
                }
            }

            var snapshot = GetCellValueSnapshot(row, column);
            if (snapshot.Kind == ExcelCellDataKind.Boolean && snapshot.Value is bool boolean) {
                return PivotFieldValue.FromBoolean(boolean);
            }

            if (snapshot.Value is double number) {
                if (TryGetPivotDateValueFromStyle(row, column, number, out var styledDate)) {
                    return PivotFieldValue.FromDate(styledDate);
                }

                return PivotFieldValue.FromNumber(number);
            }

            if (grouping?.GroupBy == GroupByValues.Range
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) {
                return PivotFieldValue.FromNumber(number);
            }

            return PivotFieldValue.FromText(text);
        }

        private bool TryGetPivotDateValueFromStyle(int row, int column, double serial, out DateTime date) {
            date = default;
            var cell = TryGetExistingCell(row, column);
            if (cell?.StyleIndex?.Value is not uint styleIndex) {
                return false;
            }

            var styles = _pivotStylesCache ??= StylesCache.Build(_spreadSheetDocument);
            if (!styles.HasDateStyles || !styles.IsDateLike(styleIndex)) {
                return false;
            }

            try {
                date = ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem);
                return true;
            } catch (ArgumentException) {
                date = default;
                return false;
            }
        }

        internal bool IsPivotDateSourceValue(int row, int column) {
            ExcelCellData snapshot = GetCellValueSnapshot(row, column);
            return snapshot.Value is double serial
                && TryGetPivotDateValueFromStyle(row, column, serial, out _);
        }

        private PivotFieldValue GetPivotFieldValue(object? value) {
            if (value == null || value == DBNull.Value) {
                return PivotFieldValue.Blank();
            }

            return value switch {
                bool boolean => PivotFieldValue.FromBoolean(boolean),
                byte number => PivotFieldValue.FromNumber(number),
                sbyte number => PivotFieldValue.FromNumber(number),
                short number => PivotFieldValue.FromNumber(number),
                ushort number => PivotFieldValue.FromNumber(number),
                int number => PivotFieldValue.FromNumber(number),
                uint number => PivotFieldValue.FromNumber(number),
                long number => PivotFieldValue.FromNumber(number),
                ulong number when number <= long.MaxValue => PivotFieldValue.FromNumber(number),
                float number => PivotFieldValue.FromNumber(number),
                double number => PivotFieldValue.FromNumber(number),
                decimal number => PivotFieldValue.FromNumber((double)number),
                DateTime dateTime => PivotFieldValue.FromDate(dateTime),
                DateTimeOffset dateTimeOffset => PivotFieldValue.FromDate(_excelDocument.DateTimeOffsetWriteStrategy(dateTimeOffset)),
#if NET6_0_OR_GREATER
                DateOnly dateOnly => PivotFieldValue.FromDate(dateOnly.ToDateTime(TimeOnly.MinValue)),
#endif
                string text => CreatePivotFieldTextValue(TrimPivotFieldText(text)),
                _ => PivotFieldValue.FromText(FormatPivotFieldText(value, _excelDocument.DateTimeOffsetWriteStrategy, _excelDocument.DateSystem))
            };
        }

        private string GetPivotFieldText(object? value) {
            if (value == null || value == DBNull.Value) {
                return string.Empty;
            }

            return FormatPivotFieldText(value, _excelDocument.DateTimeOffsetWriteStrategy, _excelDocument.DateSystem);
        }

        private static PivotFieldValue CreatePivotFieldTextValue(string text)
            => text.Length == 0 ? PivotFieldValue.Blank() : PivotFieldValue.FromText(text);

        private static string FormatPivotFieldText(object value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem) {
            return value switch {
                string text => TrimPivotFieldText(text),
                bool boolean => boolean ? "1" : "0",
                DateTime dateTime => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateTime, dateSystem)),
                DateTimeOffset dateTimeOffset => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateTimeOffsetWriteStrategy(dateTimeOffset), dateSystem)),
#if NET6_0_OR_GREATER
                DateOnly dateOnly => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateOnly.ToDateTime(TimeOnly.MinValue), dateSystem)),
#endif
                double number => InvariantNumberText.Get(number),
                float number => InvariantNumberText.Get(number),
                decimal number => number.ToString(CultureInfo.InvariantCulture),
                IFormattable formattable => TrimPivotFieldText(formattable.ToString(null, CultureInfo.InvariantCulture)),
                _ => TrimPivotFieldText(value.ToString())
            };
        }

        private static string TrimPivotFieldText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return string.Empty;
            }

            string normalized = text!;
            int last = normalized.Length - 1;
            return char.IsWhiteSpace(normalized[0]) || char.IsWhiteSpace(normalized[last])
                ? normalized.Trim()
                : normalized;
        }

        private PivotFieldValue GetGeneratedPivotDateFieldValue(int row, int column, GroupByValues groupBy) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            return TryGetPivotDateValue(row, column, text, out var date)
                ? PivotFieldValue.FromText(FormatGeneratedDateGroupValue(date, groupBy))
                : PivotFieldValue.FromText(text);
        }

        private bool TryGetPivotDateValue(int row, int column, string text, out DateTime date) {
            var snapshot = GetCellValueSnapshot(row, column);
            if (snapshot.Value is double serial) {
                try {
                    date = ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem);
                    return true;
                } catch {
                    // Fall through to string parsing when a numeric value is not a valid Excel date.
                }
            }

            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) {
                return true;
            }

            date = default;
            return false;
        }

        private static string FormatGeneratedDateGroupValue(DateTime date, GroupByValues groupBy) {
            if (groupBy == GroupByValues.Years) return date.Year.ToString(CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Quarters) return $"Q{((date.Month - 1) / 3) + 1}";
            if (groupBy == GroupByValues.Months) return date.ToString("MMMM", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Days) return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Hours) return date.Hour.ToString("00", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Minutes) return date.ToString("HH:mm", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Seconds) return date.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            return date.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string GetDateGroupFieldSuffix(GroupByValues groupBy) {
            if (groupBy == GroupByValues.Years) return "Years";
            if (groupBy == GroupByValues.Quarters) return "Quarters";
            if (groupBy == GroupByValues.Months) return "Months";
            if (groupBy == GroupByValues.Days) return "Days";
            if (groupBy == GroupByValues.Hours) return "Hours";
            if (groupBy == GroupByValues.Minutes) return "Minutes";
            if (groupBy == GroupByValues.Seconds) return "Seconds";
            return groupBy.ToString();
        }

        private static SharedItems BuildSharedItems(PivotFieldValues values, ExcelPivotGrouping? grouping, bool appendItems = true) {
            bool hasBlank = false;
            bool hasDate = false;
            bool hasNumber = false;
            bool hasString = false;
            bool containsInteger = true;
            double minNumber = 0D;
            double maxNumber = 0D;
            DateTime minDate = default;
            DateTime maxDate = default;
            int numberCount = 0;
            int dateCount = 0;

            foreach (var value in values.Items) {
                switch (value.Kind) {
                    case PivotFieldValueKind.Blank:
                        hasBlank = true;
                        break;
                    case PivotFieldValueKind.Boolean:
                        hasString = true;
                        break;
                    case PivotFieldValueKind.Number:
                        hasNumber = true;
                        if (value.Number.HasValue) {
                            double number = value.Number.Value;
                            if (numberCount == 0) {
                                minNumber = number;
                                maxNumber = number;
                            } else {
                                if (number < minNumber) minNumber = number;
                                if (number > maxNumber) maxNumber = number;
                            }

                            if (Math.Abs(number - Math.Round(number)) >= 0.0000001d) {
                                containsInteger = false;
                            }

                            numberCount++;
                        }
                        break;
                    case PivotFieldValueKind.Date:
                        hasDate = true;
                        if (value.Date.HasValue) {
                            DateTime date = value.Date.Value;
                            if (dateCount == 0) {
                                minDate = date;
                                maxDate = date;
                            } else {
                                if (date < minDate) minDate = date;
                                if (date > maxDate) maxDate = date;
                            }

                            dateCount++;
                        }
                        break;
                    default:
                        hasString = true;
                        break;
                }
            }

            var sharedItems = new SharedItems {
                ContainsString = hasString,
                ContainsBlank = hasBlank,
                ContainsDate = hasDate,
                ContainsNumber = hasNumber
            };

            if (appendItems) {
                sharedItems.Count = (uint)values.Items.Count;
            }

            if (numberCount > 0) {
                sharedItems.MinValue = minNumber;
                sharedItems.MaxValue = maxNumber;
                sharedItems.ContainsInteger = containsInteger;
            }

            if (dateCount > 0) {
                sharedItems.MinDate = minDate;
                sharedItems.MaxDate = maxDate;
            }

            if (appendItems) {
                foreach (var value in values.Items) {
                    sharedItems.Append(value.Kind switch {
                        PivotFieldValueKind.Blank => new MissingItem(),
                        PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                        PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                        PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                        _ => new StringItem { Val = value.Text }
                    });
                }
            }

            return sharedItems;
        }

        private static FieldGroup CreatePivotFieldGroup(ExcelPivotGrouping grouping, PivotFieldValues? groupItems = null, uint? baseFieldIndex = null, uint? parentFieldIndex = null) {
            var range = new RangeProperties {
                AutoStart = grouping.AutoStart,
                AutoEnd = grouping.AutoEnd,
                GroupBy = grouping.GroupBy
            };

            if (grouping.StartDate.HasValue) range.StartDate = grouping.StartDate.Value;
            if (grouping.EndDate.HasValue) range.EndDate = grouping.EndDate.Value;
            if (grouping.StartNumber.HasValue) range.StartNumber = grouping.StartNumber.Value;
            if (grouping.EndNumber.HasValue) range.EndNum = grouping.EndNumber.Value;
            if (grouping.Interval.HasValue) range.GroupInterval = grouping.Interval.Value;

            var fieldGroup = new FieldGroup(range);
            if (baseFieldIndex.HasValue) fieldGroup.Base = baseFieldIndex.Value;
            if (parentFieldIndex.HasValue) fieldGroup.ParentId = parentFieldIndex.Value;
            if (groupItems != null) {
                fieldGroup.Append(BuildGroupItems(groupItems));
            }

            return fieldGroup;
        }

        private static GroupItems BuildGroupItems(PivotFieldValues values) {
            var groupItems = new GroupItems { Count = (uint)values.Items.Count };
            foreach (var value in values.Items) {
                groupItems.Append(value.Kind switch {
                    PivotFieldValueKind.Blank => new MissingItem(),
                    PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                    PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                    PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                    _ => new StringItem { Val = value.Text }
                });
            }

            return groupItems;
        }

        private enum PivotFieldValueKind {
            Blank,
            Text,
            Boolean,
            Number,
            Date
        }

        private sealed class PivotFieldValue {
            private PivotFieldValue(PivotFieldValueKind kind, string text, bool? boolean = null, double? number = null, DateTime? date = null) {
                Kind = kind;
                Text = text;
                Boolean = boolean;
                Number = number;
                Date = date;
            }

            public PivotFieldValueKind Kind { get; }

            public string Text { get; }

            public bool? Boolean { get; }

            public double? Number { get; }

            public DateTime? Date { get; }

            public static PivotFieldValue Blank() => new(PivotFieldValueKind.Blank, string.Empty);

            public static PivotFieldValue FromText(string text) => new(PivotFieldValueKind.Text, text);

            public static PivotFieldValue FromBoolean(bool boolean) => new(PivotFieldValueKind.Boolean, boolean ? "1" : "0", boolean: boolean);

            public static PivotFieldValue FromNumber(double number) => new(PivotFieldValueKind.Number, InvariantNumberText.Get(number), number: number);

            public static PivotFieldValue FromDate(DateTime date) => new(PivotFieldValueKind.Date, date.ToString("O", CultureInfo.InvariantCulture), date: date);
        }

        private sealed class PivotFieldValues {
            public PivotFieldValues(IReadOnlyList<PivotFieldValue> items) {
                Items = items;
                TextValues = CreateTextValues(items);
            }

            public IReadOnlyList<PivotFieldValue> Items { get; }

            public IReadOnlyList<string> TextValues { get; }

            private static IReadOnlyList<string> CreateTextValues(IReadOnlyList<PivotFieldValue> items) {
                if (items.Count == 0) {
                    return Array.Empty<string>();
                }

                var textValues = new string[items.Count];
                for (int i = 0; i < items.Count; i++) {
                    textValues[i] = items[i].Text;
                }

                return textValues;
            }
        }

        private sealed class GeneratedPivotGroupingField {
            public GeneratedPivotGroupingField(int sourceIndex, int fieldIndex, int? parentFieldIndex, string fieldName, GroupByValues groupBy, ExcelPivotGrouping grouping) {
                SourceIndex = sourceIndex;
                FieldIndex = fieldIndex;
                ParentFieldIndex = parentFieldIndex;
                FieldName = fieldName;
                GroupBy = groupBy;
                Grouping = grouping;
            }

            public int SourceIndex { get; }

            public int FieldIndex { get; }

            public int? ParentFieldIndex { get; }

            public string FieldName { get; }

            public GroupByValues GroupBy { get; }

            public ExcelPivotGrouping Grouping { get; }
        }
    }
}
