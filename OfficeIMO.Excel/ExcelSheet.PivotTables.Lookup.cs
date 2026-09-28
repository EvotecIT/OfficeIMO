using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Reads a value from a saved pivot view, without refreshing its source or cache.
        /// Supports saved hierarchies, qualified date, numeric and manual text groups, default subtotals, collapsed groups and multiple measures.
        /// Unknown fields, items, measures or pivots return a typed #REF! error.
        /// </summary>
        /// <param name="pivotTableName">Saved pivot definition on this worksheet.</param>
        /// <param name="dataField">Measure caption or unique source field name.</param>
        /// <param name="items">Field criteria. Dates accept DateTime or workbook serials; error items use an ExcelCellData error, distinct from text with the same spelling.</param>
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
                || definition.RowFields?.ChildElements.Count > 257 || definition.ColumnFields?.ChildElements.Count > 257
                || definition.PageFields?.ChildElements.Count > 256
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
            PageField[] pageFields = definition.PageFields?.Elements<PageField>().ToArray() ?? Array.Empty<PageField>();
            int valuesAxisCount = rowFields.Count(f => f == -2) + columnFields.Count(f => f == -2);
            if (valuesAxisCount > 1 || (measures.Length > 1 && valuesAxisCount != 1))
                throw new NotSupportedException("Multiple pivot measures require exactly one Values axis.");
            var realFields = rowFields.Concat(columnFields).Where(f => f >= 0).ToArray();
            if (realFields.Distinct().Count() != realFields.Length
                || rowFields.Length != (definition.RowFields?.ChildElements.Count ?? 0) || columnFields.Length != (definition.ColumnFields?.ChildElements.Count ?? 0)
                || pageFields.Length != (definition.PageFields?.ChildElements.Count ?? 0)
                || rowFields.Concat(columnFields).Any(f => f < -2 || f == -1 || f >= fields.Length || (f >= 0 && f >= pivotFields.Length))
                || pageFields.Any(f => f.Field == null || f.Field.Value < 0 || f.Field.Value >= fields.Length || f.Field.Value >= pivotFields.Length)
                || realFields.Concat(pageFields.Select(f => f.Field!.Value)).Distinct().Count() != realFields.Length + pageFields.Length)
                throw new NotSupportedException("Pivot lookup requires distinct source axis and page fields.");
            var groupings = new PivotNumericGrouping?[fields.Length];
            var dateGroupings = new PivotDateGrouping?[fields.Length];
            var manualGroupings = new PivotManualGrouping?[fields.Length];
            int sourceFieldCount = part.PivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.WorksheetSource is WorksheetSource sourceRange
                && A1.TryParseRange(sourceRange.Reference?.Value ?? "", out _, out int firstSourceColumn, out _, out int lastSourceColumn)
                    ? lastSourceColumn - firstSourceColumn + 1 : fields.Length;
            foreach (int field in rowFields.Concat(columnFields).Where(f => f >= 0)) {
                if (IsDerivedDateGroup(fields[field])) dateGroupings[field] = ReadPivotDateGrouping(fields, field, sourceFieldCount);
                else if (IsDerivedManualGroup(fields[field])) manualGroupings[field] = ReadPivotManualGrouping(fields, pivotFields, field, sourceFieldCount);
                else if (fields[field].FieldGroup?.ParentId != null) continue;
                else groupings[field] = ReadPivotNumericGrouping(fields[field], field);
            }
            foreach (PageField page in pageFields) {
                int field = page.Field!.Value;
                if (fields[field].FieldGroup != null)
                    throw new NotSupportedException("Grouped page-field lookup is not qualified.");
                string name = fields[field].Name?.Value ?? "";
                if (!criteria.TryGetValue(name, out object? requested)) continue;
                if (page.Item == null) return PivotLookupReferenceError();
                Item[] items = pivotFields[field].Items?.Elements<Item>().ToArray() ?? Array.Empty<Item>();
                if (page.Item.Value >= items.Length) return PivotLookupReferenceError();
                Item selected = items[page.Item.Value];
                OpenXmlElement[] shared = fields[field].SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
                if (selected.Index == null || selected.Index.Value >= shared.Length) return PivotLookupReferenceError();
                OpenXmlElement key = shared[selected.Index.Value];
                object? expected = ReadPivotLookupSharedItem(key, _excelDocument.DateSystem);
                bool equal = key is MissingItem
                    ? requested == null || requested is string label && string.Equals(label, "(blank)", StringComparison.OrdinalIgnoreCase)
                    : PivotLookupValuesEqual(expected, key is DateTimeItem && requested is DateTime date
                        ? ExcelDateSystemConverter.ToSerial(date, _excelDocument.DateSystem) : requested);
                if (!equal)
                    return PivotLookupReferenceError();
                criteria.Remove(name);
            }
            var visibleFields = new HashSet<string>(rowFields.Concat(columnFields).Where(f => f >= 0).Select(f => fields[f].Name?.Value ?? ""), StringComparer.OrdinalIgnoreCase);
            if (criteria.Keys.Any(key => !visibleFields.Contains(key))) return PivotLookupReferenceError();
            int row = FindSavedPivotAxisItem(definition.RowItems, rowFields, fields, pivotFields, groupings, dateGroupings, manualGroupings, measure, criteria, _excelDocument.DateSystem);
            int column = FindSavedPivotAxisItem(definition.ColumnItems, columnFields, fields, pivotFields, groupings, dateGroupings, manualGroupings, measure, criteria, _excelDocument.DateSystem);
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
            PivotNumericGrouping?[] groupings, PivotDateGrouping?[] dateGroupings, PivotManualGrouping?[] manualGroupings,
            int measure, Dictionary<string, object?> criteria, ExcelDateSystem dateSystem) {
            if (axis == null || axis.ChildElements.Count == 0) throw new NotSupportedException("The pivot view has no saved axis items. Refresh or materialize the view first.");
            // One additional item accommodates the grand total for 100,000 real keys.
            if (axis.ChildElements.Count > 100_001) throw new NotSupportedException("The saved pivot axis exceeds lookup limits.");
            if ((long)axis.ChildElements.Count * Math.Max(1, axisFields.Length) > 1_000_000)
                throw new NotSupportedException("The saved pivot axis exceeds one million lookup field visits.");
            int realCount = axisFields.Count(f => f >= 0);
            int criterionDepth = 0;
            int depth = 0;
            int indexedItems = 0;
            var matches = new HashSet<uint>?[axisFields.Length];
            var collapsedMatches = new HashSet<uint>?[axisFields.Length];
            for (int axisPosition = 0; axisPosition < axisFields.Length; axisPosition++) {
                int field = axisFields[axisPosition];
                if (field < 0) continue;
                depth++;
                if (!criteria.TryGetValue(fields[field].Name?.Value ?? "", out object? expected)) continue;
                if (dateGroupings[field]?.GroupBy == ExcelPivotGroupBy.Years
                    && expected is IConvertible && expected is not string && expected is not bool
                    && expected is not DateTime) {
                    try {
                        double year = Convert.ToDouble(expected, CultureInfo.InvariantCulture);
                        if (year >= 1 && year <= 9999 && year == Math.Truncate(year))
                            expected = year.ToString(CultureInfo.InvariantCulture);
                    } catch (Exception exception) when (exception is FormatException || exception is InvalidCastException || exception is OverflowException) { }
                }
                if (groupings[field]?.TryResolveBoundary(expected, out string? groupLabel) == true)
                    expected = groupLabel;
                criterionDepth = depth;
                var savedItems = pivotFields[field].Items;
                OpenXmlElement[] sharedItems = dateGroupings[field]?.SavedItems ?? manualGroupings[field]?.SavedItems ?? groupings[field]?.SavedItems
                    ?? fields[field].SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
                if (savedItems == null || sharedItems.Length == 0) return -1;
                if (savedItems.ChildElements.Count > 100_001 || sharedItems.Length > 100_000)
                    throw new NotSupportedException("The saved pivot field exceeds lookup limits.");
                indexedItems += savedItems.ChildElements.Count + sharedItems.Length;
                if (indexedItems > 1_000_000) throw new NotSupportedException("The saved pivot criteria exceed one million indexed items.");
                var accepted = new HashSet<uint>();
                HashSet<uint>? acceptedDates = null;
                HashSet<uint>? collapsed = null;
                uint itemIndex = 0;
                foreach (var item in savedItems.ChildElements) {
                    string itemType = PivotLookupAttribute(item, "t");
                    if ((itemType.Length == 0 || itemType == "data")
                        && TryPivotLookupUnsigned(item, "x", uint.MaxValue, out uint sharedIndex) && sharedIndex < sharedItems.Length) {
                        var sharedItem = sharedItems[(int)sharedIndex];
                        string displayName = PivotLookupAttribute(item, "n");
                        bool keyEqual = sharedItem is MissingItem
                            ? expected == null || expected is string label && string.Equals(label, "(blank)", StringComparison.OrdinalIgnoreCase)
                            : PivotLookupValuesEqual(ReadPivotLookupSharedItem(sharedItem, dateSystem),
                                sharedItem is DateTimeItem && expected is DateTime date ? ExcelDateSystemConverter.ToSerial(date, dateSystem) : expected);
                        bool equal = keyEqual || displayName.Length > 0 && PivotLookupValuesEqual(displayName, expected);
                        if (equal) {
                            accepted.Add(itemIndex);
                            if (sharedItem is DateTimeItem) (acceptedDates ??= new HashSet<uint>()).Add(itemIndex);
                            string showDetails = PivotLookupAttribute(item, "sd");
                            if (showDetails == "0" || showDetails == "false") (collapsed ??= new HashSet<uint>()).Add(itemIndex);
                        }
                    }
                    itemIndex++;
                }
                // Native Excel resolves a serial to the displayed date key before an
                // equal numeric key in a mixed field. The two cache items stay distinct.
                if (acceptedDates != null) {
                    accepted = acceptedDates;
                    collapsed?.IntersectWith(acceptedDates);
                }
                matches[axisPosition] = accepted;
                collapsedMatches[axisPosition] = collapsed;
            }
            bool needsGrand = realCount > 0 && criterionDepth == 0;
            bool hasMeasures = axisFields.Contains(-2);
            int found = -1;
            int header = -1;
            int fallback = -1;
            int position = -1;
            var expanded = new uint[axisFields.Length];
            int previousPrefixLength = 0;
            foreach (var item in axis.ChildElements) {
                position++;
                string type = PivotLookupAttribute(item, "t");
                if (!TryPivotLookupUnsigned(item, "r", 0, out uint repeat) || repeat > axisFields.Length
                    || repeat > previousPrefixLength || item.ChildElements.Count > axisFields.Length - repeat)
                    throw new NotSupportedException("The saved pivot axis has an invalid repeated prefix.");
                if (repeat == 0) Array.Clear(expanded, 0, expanded.Length);
                int length = (int)repeat;
                foreach (var child in item.ChildElements) {
                    if (!TryPivotLookupUnsigned(child, "v", 0, out uint axisIndex)) return -1;
                    expanded[length++] = axisIndex;
                }
                bool data = type.Length == 0 || type == "data";
                previousPrefixLength = type == "grand" ? 0 : length;
                bool leaf = data && length == axisFields.Length;
                if (!TryPivotLookupUnsigned(item, "i", 0, out uint itemMeasure)) return -1;
                if (hasMeasures && itemMeasure != measure) continue;
                bool primary = true;
                bool headerTotal = false;
                if (needsGrand) {
                    if (type != "grand") continue;
                } else {
                    int representedDepth = 0;
                    for (int index = 0; index < length; index++) if (axisFields[index] >= 0) representedDepth++;
                    primary = criterionDepth == realCount ? leaf : type == "default" && representedDepth == criterionDepth;
                    if (data && !leaf && length > 0 && representedDepth == criterionDepth && axisFields[length - 1] >= 0
                        && (!hasMeasures || Array.IndexOf(axisFields, -2) < length)) {
                        var field = pivotFields[axisFields[length - 1]];
                        headerTotal = collapsedMatches[length - 1]?.Contains(expanded[length - 1]) == true
                            || (field.Outline?.Value != false && field.SubtotalTop?.Value != false && field.DefaultSubtotal?.Value != false);
                    }
                    if (!primary && !leaf && !headerTotal) continue;
                    bool accepted = true;
                    for (int index = 0; index < matches.Length; index++) {
                        if (matches[index] != null && (index >= length || !matches[index]!.Contains(expanded[index]))) { accepted = false; break; }
                    }
                    if (!accepted) continue;
                }
                // Prefer the explicit subtotal, then a total in an outline/compact
                // group header. With no subtotal, Excel accepts a
                // partial criterion only when exactly one displayed leaf matches.
                if (primary) found = found == -1 ? position : -2;
                else if (headerTotal) header = header == -1 ? position : -2;
                else fallback = fallback == -1 ? position : -2;
            }
            if (found != -1) return found >= 0 ? found : -1;
            if (header != -1) return header >= 0 ? header : -1;
            return fallback >= 0 ? fallback : -1;
        }

        private static object? ReadPivotLookupSharedItem(OpenXmlElement item, ExcelDateSystem dateSystem) => item switch {
            StringItem text => text.Val?.Value,
            NumberItem number => number.Val?.Value,
            BooleanItem boolean => boolean.Val?.Value,
            DateTimeItem date => date.Val == null ? null : ExcelPivotCacheDateCodec.ToSerial(date.Val.Value, dateSystem),
            ErrorItem error => new ExcelCellData(ExcelCellDataKind.Error, error.Val?.Value, cachedText: error.Val?.Value),
            MissingItem => null,
            _ => throw new NotSupportedException("The pivot key type is not supported.")
        };

        private static bool PivotLookupValuesEqual(object? left, object? right) {
            if (left is ExcelCellData { Kind: ExcelCellDataKind.Error } leftError
                && right is ExcelCellData { Kind: ExcelCellDataKind.Error } rightError)
                return Equals(leftError.Value, rightError.Value);
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
