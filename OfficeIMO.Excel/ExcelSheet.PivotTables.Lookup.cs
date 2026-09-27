using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Reads a value from a saved pivot view, without refreshing its source or cache.
        /// Supports ungrouped views with at most one row and one column field, including multiple measures.
        /// Unknown fields, items, measures or pivots return a typed #REF! error.
        /// </summary>
        /// <exception cref="NotSupportedException">The view has no materialized items or uses an unsupported layout or exceeds lookup limits.</exception>
        public ExcelCellData GetPivotData(string pivotTableName, string dataField, IReadOnlyDictionary<string, object?>? items = null) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) throw new ArgumentException("A pivot table name is required.", nameof(pivotTableName));
            if (string.IsNullOrWhiteSpace(dataField)) throw new ArgumentException("A data field is required.", nameof(dataField));
            return _excelDocument.ExecuteReadAfterMaterializing(() => {
                var criteria = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                if (items != null) {
                    if (items.Count > 256) throw new NotSupportedException("Pivot lookup supports at most 256 criteria.");
                    foreach (var item in items) {
                        if (string.IsNullOrWhiteSpace(item.Key)) throw new ArgumentException("Pivot item field names cannot be empty.", nameof(items));
                        if (criteria.TryGetValue(item.Key, out object? previous) && !PivotLookupValuesEqual(previous, item.Value)) return PivotLookupReferenceError();
                        criteria[item.Key] = item.Value;
                    }
                }
                var part = _worksheetPart.PivotTableParts.FirstOrDefault(p => string.Equals(p.PivotTableDefinition?.Name?.Value, pivotTableName, StringComparison.OrdinalIgnoreCase));
                return part == null ? PivotLookupReferenceError() : ReadSavedPivotData(part, dataField, criteria);
            });
        }

        private static ExcelCellData PivotLookupReferenceError() => new(ExcelCellDataKind.Error, "#REF!", cachedText: "#REF!");

        private ExcelCellData ReadSavedPivotData(PivotTablePart part, string dataField, Dictionary<string, object?> criteria) {
            var definition = part.PivotTableDefinition;
            var location = definition?.Location;
            var cacheFields = part.PivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheFields;
            if (definition == null || location == null || cacheFields == null
                || location.FirstDataRow == null || location.FirstDataColumn == null
                || !A1.TryParseRange(location.Reference?.Value ?? "", out int top, out int left, out int bottom, out int right))
                return PivotLookupReferenceError();
            if (cacheFields.ChildElements.Count > 256 || definition.PivotFields?.ChildElements.Count > 256
                || definition.DataFields?.ChildElements.Count > 256 || criteria.Count > 256
                || definition.RowFields?.ChildElements.Count > 2 || definition.ColumnFields?.ChildElements.Count > 2
                || (long)(bottom - top + 1) * (right - left + 1) > 1_000_000)
                throw new NotSupportedException("The saved pivot exceeds lookup limits.");
            var fields = cacheFields.Elements<CacheField>().ToArray();
            var pivotFields = definition.PivotFields?.Elements<PivotField>().ToArray() ?? Array.Empty<PivotField>();
            var measures = definition.DataFields?.Elements<DataField>().ToArray() ?? Array.Empty<DataField>();
            int measure = -1;
            for (int index = 0; index < measures.Length; index++) {
                if (!string.Equals(measures[index].Name?.Value, dataField, StringComparison.OrdinalIgnoreCase)) continue;
                if (measure >= 0) return PivotLookupReferenceError();
                measure = index;
            }
            if (measure < 0) {
                for (int index = 0; index < measures.Length; index++) {
                    uint? source = measures[index].Field?.Value;
                    if (!source.HasValue || source.Value >= fields.Length
                        || !string.Equals(fields[source.Value].Name?.Value, dataField, StringComparison.OrdinalIgnoreCase)) continue;
                    if (measure >= 0) return PivotLookupReferenceError();
                    measure = index;
                }
            }
            if (measure < 0) return PivotLookupReferenceError();
            int[] rowFields = definition.RowFields?.Elements<Field>().Select(f => f.Index?.Value ?? int.MinValue).ToArray() ?? Array.Empty<int>();
            int[] columnFields = definition.ColumnFields?.Elements<Field>().Select(f => f.Index?.Value ?? int.MinValue).ToArray() ?? Array.Empty<int>();
            int valuesAxisCount = rowFields.Count(f => f == -2) + columnFields.Count(f => f == -2);
            if (valuesAxisCount > 1 || (measures.Length > 1 && valuesAxisCount != 1))
                throw new NotSupportedException("Multiple pivot measures require exactly one Values axis.");
            if (rowFields.Length > 2 || columnFields.Length > 2 || rowFields.Count(f => f >= 0) > 1 || columnFields.Count(f => f >= 0) > 1
                || definition.PageFields?.ChildElements.Count > 0 || rowFields.Concat(columnFields).Any(f => f < -2 || f == -1 || f >= fields.Length || (f >= 0 && f >= pivotFields.Length)))
                throw new NotSupportedException("Pivot lookup requires an ungrouped view with at most one row and one column field and no page fields.");
            foreach (int field in rowFields.Concat(columnFields).Where(f => f >= 0)) {
                if (fields[field].FieldGroup != null) throw new NotSupportedException("Grouped pivot lookup is not supported.");
            }
            var visibleFields = new HashSet<string>(rowFields.Concat(columnFields).Where(f => f >= 0).Select(f => fields[f].Name?.Value ?? ""), StringComparer.OrdinalIgnoreCase);
            if (criteria.Keys.Any(key => !visibleFields.Contains(key))) return PivotLookupReferenceError();
            int row = FindSavedPivotAxisItem(definition.RowItems, rowFields, fields, pivotFields, measure, criteria);
            int column = FindSavedPivotAxisItem(definition.ColumnItems, columnFields, fields, pivotFields, measure, criteria);
            if (row < 0 || column < 0) return PivotLookupReferenceError();
            long outputRow = (long)top + location.FirstDataRow.Value + row;
            long outputColumn = (long)left + location.FirstDataColumn.Value + column;
            if (outputRow > bottom || outputColumn > right || outputRow < top || outputColumn < left) return PivotLookupReferenceError();
            Cell? cell = TryGetExistingCell((int)outputRow, (int)outputColumn);
            var value = GetCellValueSnapshot(cell);
            if (value.Kind == ExcelCellDataKind.Blank) return new ExcelCellData(ExcelCellDataKind.Number, 0d);
            if (cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error)
                return new ExcelCellData(ExcelCellDataKind.Error, value.Value, cachedText: value.CachedText);
            return value;
        }

        private static int FindSavedPivotAxisItem(OpenXmlCompositeElement? axis, int[] axisFields, CacheField[] fields, PivotField[] pivotFields,
            int measure, Dictionary<string, object?> criteria) {
            if (axis == null || axis.ChildElements.Count == 0) throw new NotSupportedException("The pivot view has no saved axis items. Refresh or materialize the view first.");
            // One additional item accommodates the grand total for 100,000 real keys.
            if (axis.ChildElements.Count > 100_001) throw new NotSupportedException("The saved pivot axis exceeds lookup limits.");
            int realField = -1;
            foreach (int field in axisFields) if (field >= 0) { realField = field; break; }
            bool hasCriterion = realField >= 0 && criteria.TryGetValue(fields[realField].Name?.Value ?? "", out _);
            bool needsGrand = realField >= 0 && !hasCriterion;
            bool hasMeasures = axisFields.Contains(-2);
            OpenXmlElement[] fieldItems = Array.Empty<OpenXmlElement>();
            OpenXmlElement[] sharedItems = Array.Empty<OpenXmlElement>();
            if (hasCriterion) {
                var savedItems = pivotFields[realField].Items;
                var shared = fields[realField].SharedItems;
                if (savedItems == null || shared == null) return -1;
                // Pivot fields also carry a default subtotal item, even when totals are hidden.
                if (savedItems.ChildElements.Count > 100_001 || shared.ChildElements.Count > 100_000)
                    throw new NotSupportedException("The saved pivot field exceeds lookup limits.");
                fieldItems = savedItems.ChildElements.ToArray();
                sharedItems = shared.ChildElements.ToArray();
            }
            int found = -1;
            int position = -1;
            foreach (var item in axis.ChildElements) {
                position++;
                string type = PivotLookupAttribute(item, "t");
                if (!TryPivotLookupUnsigned(item, "r", 0, out uint repeat) || repeat != 0)
                    throw new NotSupportedException("Compressed pivot axis prefixes are not supported.");
                if (!TryPivotLookupUnsigned(item, "i", 0, out uint itemMeasure)) return -1;
                if (hasMeasures && itemMeasure != measure) continue;
                if (needsGrand ? type != "grand" : type.Length != 0 && type != "data") continue;
                if (hasCriterion) {
                    var indices = item.ChildElements;
                    int realPosition = Array.IndexOf(axisFields, realField);
                    if (indices.Count <= realPosition || !TryPivotLookupUnsigned(indices[realPosition], "v", 0, out uint key)) return -1;
                    if (key >= fieldItems.Length) return -1;
                    if (!TryPivotLookupUnsigned(fieldItems[(int)key], "x", uint.MaxValue, out uint sharedIndex) || sharedIndex >= sharedItems.Length) return -1;
                    object? expected = criteria[fields[realField].Name?.Value ?? ""];
                    var sharedItem = sharedItems[(int)sharedIndex];
                    bool matches = sharedItem is MissingItem
                        ? expected == null || expected is string label && string.Equals(label, "(blank)", StringComparison.OrdinalIgnoreCase)
                        : PivotLookupValuesEqual(ReadPivotLookupSharedItem(sharedItem), expected);
                    if (!matches) continue;
                }
                if (found >= 0) return -1;
                found = position;
            }
            return found;
        }

        private static object? ReadPivotLookupSharedItem(OpenXmlElement item) => item switch {
            StringItem text => text.Val?.Value,
            NumberItem number => number.Val?.Value,
            BooleanItem boolean => boolean.Val?.Value,
            MissingItem => null,
            _ => throw new NotSupportedException("Pivot lookup supports text, numeric, Boolean and blank item keys.")
        };

        private static bool PivotLookupValuesEqual(object? left, object? right) {
            if (left is string leftText && right is string rightText) return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
            if (left is double number && right is IConvertible && right is not bool && right is not string && right is not DateTime) {
                try { return number == Convert.ToDouble(right, CultureInfo.InvariantCulture); }
                catch (Exception exception) when (exception is FormatException || exception is InvalidCastException || exception is OverflowException) { return false; }
            }
            return Equals(left, right);
        }

        private static string PivotLookupAttribute(OpenXmlElement element, string name) {
            foreach (var attribute in element.GetAttributes())
                if (attribute.LocalName == name && attribute.NamespaceUri.Length == 0) return attribute.Value ?? "";
            return "";
        }
        private static bool TryPivotLookupUnsigned(OpenXmlElement element, string name, uint fallback, out uint value) {
            string text = PivotLookupAttribute(element, name);
            value = fallback;
            return text.Length == 0 || uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value);
        }
    }
}
