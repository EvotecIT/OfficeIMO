using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const long MaximumPivotFieldScanCells = 1_000_000L;

        private static void ValidatePivotFieldScanBudget(int firstDataRow, int lastDataRow, IReadOnlyList<bool> collectFieldValues, int generatedFieldCount) {
            long rowCount = Math.Max(0L, (long)lastDataRow - firstDataRow + 1L);
            int collectedFields = generatedFieldCount;
            for (int field = 0; field < collectFieldValues.Count; field++) {
                if (collectFieldValues[field]) {
                    collectedFields++;
                }
            }

            long scanCells = rowCount * collectedFields;
            if (scanCells > MaximumPivotFieldScanCells) {
                throw new InvalidOperationException($"Pivot shared-item discovery would inspect {scanCells} cells, exceeding the supported limit of {MaximumPivotFieldScanCells}.");
            }
        }

        private List<PivotFieldValues> BuildPivotFieldValueMap(int fieldCount, int firstDataRow, int lastDataRow, int firstColumn,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap, IReadOnlyList<bool>? collectFieldValues = null) {
            var fieldValues = new List<PivotFieldValue>[fieldCount];
            var seenValues = new HashSet<PivotFieldValue>[fieldCount];
            var fieldGroupings = new ExcelPivotGrouping?[fieldCount];
            var fieldsToCollect = new List<int>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                fieldValues[field] = new List<PivotFieldValue>();
                groupingMap.TryGetValue(field, out fieldGroupings[field]);
                bool collectValues = ShouldCollectPivotSharedItems(field, collectFieldValues);
                if (collectValues) {
                    seenValues[field] = new HashSet<PivotFieldValue>();
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
                        if (seenValues[field]!.Add(value)) {
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
            var seenValues = new HashSet<PivotFieldValue>[fieldCount];
            var fieldsToCollect = new List<int>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                fieldValues[field] = new List<PivotFieldValue>();
                if (ShouldCollectPivotSharedItems(field, collectFieldValues)) {
                    seenValues[field] = new HashSet<PivotFieldValue>();
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
                        if (seenValues[field]!.Add(value)) {
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

        private List<PivotFieldValues> BuildGeneratedPivotFieldValueMap(
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            int firstDataRow,
            int lastDataRow,
            int firstColumn) {
            var maps = new List<PivotFieldValues>(generatedFields.Count);
            foreach (var generatedField in generatedFields) {
                var values = new List<PivotFieldValue>();
                var seen = new HashSet<PivotFieldValue>();
                int column = firstColumn + generatedField.SourceIndex;
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    var value = GetGeneratedPivotDateFieldValue(row, column, generatedField.GroupBy);
                    if (seen.Add(value)) {
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

        private static Dictionary<PivotFieldValue, uint>?[] BuildPivotRecordSharedItemLookups(
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<bool>? sharedItemFields) {
            var lookups = new Dictionary<PivotFieldValue, uint>?[fieldValueMap.Count];
            for (int field = 0; field < fieldValueMap.Count; field++) {
                if (sharedItemFields != null
                    && (field < 0 || field >= sharedItemFields.Count || !sharedItemFields[field])) {
                    continue;
                }

                var items = fieldValueMap[field].Items;
                if (items.Count == 0) {
                    continue;
                }

                var lookup = new Dictionary<PivotFieldValue, uint>();
                for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
                    lookup[items[itemIndex]] = (uint)itemIndex;
                }

                lookups[field] = lookup;
            }

            return lookups;
        }

        private void AppendGeneratedPivotCacheRecordItems(
            PivotCacheRecord record,
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<Dictionary<PivotFieldValue, uint>?> generatedLookups,
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
            IReadOnlyList<Dictionary<PivotFieldValue, uint>?> generatedLookups,
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
            IReadOnlyList<Dictionary<PivotFieldValue, uint>?> generatedLookups,
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

        private static void WritePivotCacheRecordItemXml(TextWriter writer, PivotFieldValue value, IReadOnlyDictionary<PivotFieldValue, uint>? sharedItems) {
            if (sharedItems != null && sharedItems.TryGetValue(value, out uint index)) {
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
                case PivotFieldValueKind.Error:
                    writer.Write("<e v=\"");
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

        private static OpenXmlElement CreatePivotCacheRecordItem(PivotFieldValue value, IReadOnlyDictionary<PivotFieldValue, uint>? sharedItems) {
            if (sharedItems != null && sharedItems.TryGetValue(value, out uint index)) {
                return new FieldItem { Val = index };
            }

            return value.Kind switch {
                PivotFieldValueKind.Blank => new MissingItem(),
                PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                PivotFieldValueKind.Error => new ErrorItem { Val = value.Text },
                _ => new StringItem { Val = value.Text }
            };
        }

    }
}
