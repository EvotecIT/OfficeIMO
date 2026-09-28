using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.RegularExpressions;
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

            foreach (var filter in filters.Where(IsMaterializedPivotLabelComparison)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date
                    || key.Kind == PivotFieldValueKind.Number))
                    throw new NotSupportedException("Label-filter materialization has not qualified formatted date or numeric captions.");
                string needle = filter.StringValue1?.Value
                    ?? throw new NotSupportedException("The label filter has no saved criterion.");
                var type = filter.Type!.Value;
                HashSet<PivotFieldValue> included;
                if (IsMaterializedPivotLabelRange(type)) {
                    string? second = filter.StringValue2?.Value;
                    ValidateMaterializedPivotLabelRangePredicate(filter, needle, second);
                    if (!IsQualifiedPivotLabelOrderingText(needle)
                        || (second != null && !IsQualifiedPivotLabelOrderingText(second))
                        || maps[field].Items.Any(key => {
                            string label = captions.TryGetValue(field, out var names) && names.TryGetValue(key, out string? caption)
                                ? caption : PivotMaterializedCaption(key, _excelDocument.DateSystem);
                            return !IsQualifiedPivotLabelOrderingText(label);
                        }))
                        throw new NotSupportedException("Label range materialization requires ASCII alphabetic captions and criteria; localized text ordering is not qualified.");
                    included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                        string label = captions.TryGetValue(field, out var names) && names.TryGetValue(key, out string? caption)
                            ? caption : PivotMaterializedCaption(key, _excelDocument.DateSystem);
                        return MatchesMaterializedPivotLabelRange(label, needle, second, type);
                    }));
                } else {
                    string savedPattern = NormalizePivotFilterAutoFilterValue(type, needle);
                    var comparison = ResolveSingleFilterOperator(type);
                    ValidateMaterializedPivotLabelPredicate(filter, savedPattern, comparison);
                    var pattern = new Regex("\\A" + CreateFormulaWildcardPattern(savedPattern) + "\\z",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, FormulaRegexTimeout);
                    bool negate = comparison == FilterOperatorValues.NotEqual;
                    included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                        string label = captions.TryGetValue(field, out var names) && names.TryGetValue(key, out string? caption)
                            ? caption : PivotMaterializedCaption(key, _excelDocument.DateSystem);
                        return pattern.IsMatch(label) != negate;
                    }));
                }
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(IsMaterializedPivotValueComparison)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (axisFields.Count != 1 || filter.MeasureField == null || filter.MeasureField.Value >= measures.Length)
                    throw new NotSupportedException("Value-filter materialization requires one ordinary axis field and a saved measure.");
                var custom = QualifiedPivotCustomFilter(filter);
                var comparison = ResolveSingleFilterOperator(filter.Type!.Value);
                if ((custom.Operator?.Value ?? FilterOperatorValues.Equal) != comparison
                    || !double.TryParse(custom.Val?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double threshold)
                    || double.IsNaN(threshold) || double.IsInfinity(threshold))
                    throw new NotSupportedException("The value filter does not have a qualified finite comparison threshold.");
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
                    pair.Value.GetValue(function).Value is double value
                    && MatchesMaterializedPivotValueComparison(value, threshold, comparison)).Select(pair => pair.Key));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            if (filters.Any(filter => !IsMaterializedPivotLabelComparison(filter)
                && !IsMaterializedPivotValueComparison(filter)))
                throw new NotSupportedException("This pivot label or value filter is not qualified for materialization.");
            if (!visibility.IncludedRows.Any(include => include))
                throw new NotSupportedException("The pivot filters select no source records for materialization.");
        }

        private static bool IsMaterializedPivotLabelComparison(PivotFilter filter) {
            var type = filter.Type?.Value;
            return type == PivotFilterValues.CaptionEqual || type == PivotFilterValues.CaptionNotEqual
                || type == PivotFilterValues.CaptionBeginsWith || type == PivotFilterValues.CaptionNotBeginsWith
                || type == PivotFilterValues.CaptionEndsWith || type == PivotFilterValues.CaptionNotEndsWith
                || type == PivotFilterValues.CaptionContains || type == PivotFilterValues.CaptionNotContains
                || type == PivotFilterValues.CaptionGreaterThan || type == PivotFilterValues.CaptionGreaterThanOrEqual
                || type == PivotFilterValues.CaptionLessThan || type == PivotFilterValues.CaptionLessThanOrEqual
                || type == PivotFilterValues.CaptionBetween || type == PivotFilterValues.CaptionNotBetween;
        }

        private static bool IsMaterializedPivotLabelRange(PivotFilterValues type)
            => type == PivotFilterValues.CaptionGreaterThan || type == PivotFilterValues.CaptionGreaterThanOrEqual
                || type == PivotFilterValues.CaptionLessThan || type == PivotFilterValues.CaptionLessThanOrEqual
                || type == PivotFilterValues.CaptionBetween || type == PivotFilterValues.CaptionNotBetween;

        private static bool IsQualifiedPivotLabelOrderingText(string text)
            => text.Length > 0 && text.All(character => character is >= 'A' and <= 'Z' or >= 'a' and <= 'z');

        private static void ValidateMaterializedPivotLabelRangePredicate(PivotFilter filter, string first, string? second) {
            PivotFilterValues type = filter.Type!.Value;
            if (!TryResolveBetweenFilter(type, out var firstOperator, out var secondOperator, out bool matchAll)) {
                ValidateMaterializedPivotLabelPredicate(filter, first, ResolveSingleFilterOperator(type));
                return;
            }
            if (second == null)
                throw new NotSupportedException("The label range filter has no saved second criterion.");
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (columns.Length != 1 || columns[0].ColumnId?.Value != 0
                || columns[0].ChildElements.Count != 1
                || columns[0].GetFirstChild<CustomFilters>() is not CustomFilters custom
                || (custom.And?.Value == true) != matchAll
                || custom.ChildElements.Count != 2)
                throw new NotSupportedException("The label range filter does not have a qualified saved predicate.");
            var predicates = custom.Elements<CustomFilter>().ToArray();
            if (predicates.Length != 2 || (predicates[0].Operator?.Value ?? FilterOperatorValues.Equal) != firstOperator
                || predicates[0].Val?.Value != first || (predicates[1].Operator?.Value ?? FilterOperatorValues.Equal) != secondOperator
                || predicates[1].Val?.Value != second)
                throw new NotSupportedException("The label range filter does not have a qualified saved predicate.");
        }

        private static bool MatchesMaterializedPivotLabelRange(string label, string first, string? second,
            PivotFilterValues type) {
            int firstComparison = StringComparer.OrdinalIgnoreCase.Compare(label, first);
            if (type == PivotFilterValues.CaptionGreaterThan) return firstComparison > 0;
            if (type == PivotFilterValues.CaptionGreaterThanOrEqual) return firstComparison >= 0;
            if (type == PivotFilterValues.CaptionLessThan) return firstComparison < 0;
            if (type == PivotFilterValues.CaptionLessThanOrEqual) return firstComparison <= 0;
            int secondComparison = StringComparer.OrdinalIgnoreCase.Compare(label, second);
            if (type == PivotFilterValues.CaptionBetween) return firstComparison >= 0 && secondComparison <= 0;
            return firstComparison < 0 || secondComparison > 0;
        }

        private static void ValidateMaterializedPivotLabelPredicate(PivotFilter filter, string pattern,
            FilterOperatorValues comparison) {
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (comparison == FilterOperatorValues.Equal && filter.Type?.Value == PivotFilterValues.CaptionEqual
                && columns.Length == 1 && columns[0].ColumnId?.Value == 0
                && columns[0].ChildElements.Count == 1
                && columns[0].GetFirstChild<Filters>() is Filters values
                && values.ChildElements.Count == 1 && values.GetFirstChild<Filter>()?.Val?.Value == pattern)
                return;
            var custom = QualifiedPivotCustomFilter(filter);
            if ((custom.Operator?.Value ?? FilterOperatorValues.Equal) != comparison
                || !string.Equals(custom.Val?.Value, pattern, StringComparison.Ordinal))
                throw new NotSupportedException("The label filter does not have a qualified saved predicate.");
        }

        private static bool IsMaterializedPivotValueComparison(PivotFilter filter) {
            var type = filter.Type?.Value;
            return type == PivotFilterValues.ValueEqual || type == PivotFilterValues.ValueNotEqual
                || type == PivotFilterValues.ValueGreaterThan || type == PivotFilterValues.ValueGreaterThanOrEqual
                || type == PivotFilterValues.ValueLessThan || type == PivotFilterValues.ValueLessThanOrEqual;
        }

        private static bool MatchesMaterializedPivotValueComparison(double value, double threshold,
            FilterOperatorValues comparison) {
            if (comparison == FilterOperatorValues.Equal) return value == threshold;
            if (comparison == FilterOperatorValues.NotEqual) return value != threshold;
            if (comparison == FilterOperatorValues.GreaterThan) return value > threshold;
            if (comparison == FilterOperatorValues.GreaterThanOrEqual) return value >= threshold;
            if (comparison == FilterOperatorValues.LessThan) return value < threshold;
            if (comparison == FilterOperatorValues.LessThanOrEqual) return value <= threshold;
            return false;
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
