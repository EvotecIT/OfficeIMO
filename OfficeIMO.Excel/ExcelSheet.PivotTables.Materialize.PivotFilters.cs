using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void ApplyMaterializedPivotFilters(
            ExcelSheet source, PivotTableDefinition definition, CacheField[] cacheFields,
            IReadOnlyList<PivotFieldValues> maps,
            IReadOnlyDictionary<int, Dictionary<PivotFieldValue, string>> captions,
            DataField[] measures, IReadOnlyList<int> axisFields,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings,
            IReadOnlyDictionary<int, PivotManualGrouping> manualGroupings,
            PivotMaterializationVisibility visibility,
            int firstRow, int lastRow, int firstColumn, int limit, CancellationToken token) {
            if (definition.PivotFilters == null) return;
            var filters = definition.PivotFilters.Elements<PivotFilter>().ToArray();
            if (filters.Length != definition.PivotFilters.ChildElements.Count || filters.Length > 256
                || (long)(lastRow - firstRow) * filters.Length > limit)
                throw new NotSupportedException("The pivot filters exceed the qualified materialization budget or shape.");

            foreach (var filter in filters.Where(filter => filter.Type?.Value == PivotFilterValues.CaptionContains)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                string needle = filter.StringValue1?.Value
                    ?? throw new NotSupportedException("The label filter has no saved criterion.");
                var custom = QualifiedPivotCustomFilter(filter);
                if (custom.Operator != null && custom.Operator.Value != FilterOperatorValues.Equal
                    || !string.Equals(custom.Val?.Value, "*" + needle + "*", StringComparison.Ordinal))
                    throw new NotSupportedException("The label filter does not have a qualified contains predicate.");
                var included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                    string label = captions.TryGetValue(field, out var names) && names.TryGetValue(key, out string? caption)
                        ? caption : key.Text;
                    return label.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
                }));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(filter => filter.Type?.Value == PivotFilterValues.ValueGreaterThan)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (axisFields.Count != 1 || filter.MeasureField == null || filter.MeasureField.Value >= measures.Length)
                    throw new NotSupportedException("Value-filter materialization requires one ordinary axis field and a saved measure.");
                var custom = QualifiedPivotCustomFilter(filter);
                if (custom.Operator?.Value != FilterOperatorValues.GreaterThan
                    || !double.TryParse(custom.Val?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double threshold)
                    || double.IsNaN(threshold) || double.IsInfinity(threshold))
                    throw new NotSupportedException("The value filter does not have a finite greater-than threshold.");
                var measure = measures[filter.MeasureField.Value];
                int measureColumn = firstColumn + (int)measure.Field!.Value;
                var aggregates = new Dictionary<PivotFieldValue, ExcelPivotAggregateAccumulator>();
                for (int row = firstRow + 1; row <= lastRow; row++) {
                    token.ThrowIfCancellationRequested();
                    if (!visibility.IncludedRows[row - firstRow - 1]) continue;
                    var key = MaterializedPivotAxisKey(source, row, firstColumn + field, field,
                        groupings, dateGroupings, manualGroupings);
                    if (!aggregates.TryGetValue(key, out var aggregate)) {
                        if (aggregates.Count >= 100_000 || aggregates.Count >= limit)
                            throw new InvalidOperationException("The value filter exceeds the pivot aggregate budget.");
                        aggregate = new ExcelPivotAggregateAccumulator();
                        aggregates.Add(key, aggregate);
                    }
                    var cell = source.TryGetExistingCell(row, measureColumn);
                    aggregate.Add(source.GetCellValueSnapshot(cell).Value,
                        cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error);
                }
                var function = (measure.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum).ToOfficeEnum();
                var included = new HashSet<PivotFieldValue>(aggregates.Where(pair =>
                    pair.Value.GetValue(function).Value is double value && value > threshold).Select(pair => pair.Key));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            if (filters.Any(filter => filter.Type?.Value != PivotFilterValues.CaptionContains
                && filter.Type?.Value != PivotFilterValues.ValueGreaterThan))
                throw new NotSupportedException("This pivot label or value filter is not qualified for materialization.");
            if (!visibility.IncludedRows.Any(include => include))
                throw new NotSupportedException("The pivot filters select no source records for materialization.");
        }

        private static int QualifiedPivotFilterField(PivotFilter filter, CacheField[] fields, IReadOnlyList<int> axisFields) {
            if (filter.Field == null || filter.Field.Value >= fields.Length
                || !axisFields.Contains((int)filter.Field.Value)
                || fields[filter.Field.Value].DatabaseField?.Value == false
                || fields[filter.Field.Value].FieldGroup != null)
                throw new NotSupportedException("The pivot filter requires an ordinary row or column source field.");
            return (int)filter.Field.Value;
        }

        private static CustomFilter QualifiedPivotCustomFilter(PivotFilter filter) {
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (columns.Length != 1 || columns[0].ColumnId?.Value != 0
                || columns[0].ChildElements.Count != 1
                || columns[0].GetFirstChild<CustomFilters>() is not CustomFilters customFilters
                || customFilters.Elements<CustomFilter>().Count() != 1)
                throw new NotSupportedException("The pivot filter requires one saved custom-filter predicate.");
            return customFilters.Elements<CustomFilter>().Single();
        }

        private static void FilterMaterializedSourceRows(
            ExcelSheet source, bool[] includedRows, HashSet<PivotFieldValue> includedKeys, int field,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings,
            IReadOnlyDictionary<int, PivotManualGrouping> manualGroupings,
            int firstRow, int lastRow, int firstColumn, CancellationToken token) {
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                int index = row - firstRow - 1;
                if (!includedRows[index]) continue;
                var key = MaterializedPivotAxisKey(source, row, firstColumn + field, field,
                    groupings, dateGroupings, manualGroupings);
                if (!includedKeys.Contains(key)) includedRows[index] = false;
            }
        }
    }
}
