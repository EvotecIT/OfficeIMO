using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

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
            int firstRow, int lastRow, int firstColumn, int limit,
            DateTime referenceDate, CancellationToken token) {
            if (definition.PivotFilters == null) return;
            var filters = definition.PivotFilters.Elements<PivotFilter>().ToArray();
            if (filters.Length != definition.PivotFilters.ChildElements.Count || filters.Length > 256
                || (long)(lastRow - firstRow) * filters.Length > limit)
                throw new NotSupportedException("The pivot filters exceed the qualified materialization budget or shape.");

            foreach (var filter in filters.Where(IsMaterializedPivotFixedDateFilter)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (!maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date)
                    || maps[field].Items.Any(key => key.Kind != PivotFieldValueKind.Date && key.Kind != PivotFieldValueKind.Blank))
                    throw new NotSupportedException("Fixed-date filter materialization requires a date-valued pivot field.");
                PivotFilterValues type = filter.Type!.Value;
                bool wholeDay = QualifiedPivotWholeDay(filter);
                double first, second = 0;
                if (type == PivotFilterValues.DateBetween || type == PivotFilterValues.DateNotBetween) {
                    (first, second) = QualifiedPivotValueRange(filter);
                } else {
                    var predicate = QualifiedPivotCustomFilter(filter);
                    if ((predicate.Operator?.Value ?? FilterOperatorValues.Equal) != ResolveSingleFilterOperator(type)
                        || !TryFinitePivotThreshold(predicate.Val?.Value, out first)
                        || (filter.StringValue1 != null && (!TryFinitePivotThreshold(filter.StringValue1.Value, out double savedFirst)
                            || savedFirst != first)) || filter.StringValue2 != null)
                        throw new NotSupportedException("The fixed-date filter has no qualified saved predicate.");
                }
                double lower = wholeDay ? Math.Floor(first) : first;
                double upper = wholeDay ? Math.Floor(second) : second;
                var included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                    if (key.Kind != PivotFieldValueKind.Date) return false;
                    double serial = ExcelPivotCacheDateCodec.ToSerial(key.Date!.Value, _excelDocument.DateSystem);
                    if (wholeDay) {
                        serial = Math.Floor(serial);
                    }
                    return type == PivotFilterValues.DateBetween ? serial >= lower && serial <= upper
                        : type == PivotFilterValues.DateNotBetween ? serial < lower || serial > upper
                        : MatchesMaterializedPivotValueComparison(serial, lower, ResolveSingleFilterOperator(type));
                }));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(IsMaterializedPivotCalendarFilter)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (!maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date)
                    || maps[field].Items.Any(key => key.Kind != PivotFieldValueKind.Date && key.Kind != PivotFieldValueKind.Blank))
                    throw new NotSupportedException("Calendar-period filter materialization requires a date-valued pivot field.");
                PivotFilterValues type = filter.Type!.Value;
                if (!TryMaterializedPivotCalendarMonths(type, out int firstMonth, out int lastMonth)
                    || filter.MeasureField != null || filter.StringValue1 != null || filter.StringValue2 != null)
                    throw new NotSupportedException("The calendar-period filter has no qualified saved predicate.");
                var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
                if (filter.AutoFilter?.ChildElements.Count != 1 || columns.Length != 1
                    || columns[0].ColumnId?.Value != 0 || columns[0].ChildElements.Count != 1
                    || columns[0].GetFirstChild<DynamicFilter>() is not DynamicFilter dynamic
                    || dynamic.ChildElements.Count != 0 || dynamic.GetAttributes().Count != 1
                    || dynamic.Type?.Value != ResolveDynamicFilterType(type))
                    throw new NotSupportedException("The calendar-period filter has no qualified saved dynamic predicate.");
                var included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key =>
                    key.Kind == PivotFieldValueKind.Date
                    && key.Date!.Value.Month >= firstMonth && key.Date.Value.Month <= lastMonth));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(IsMaterializedPivotRelativeDateFilter)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (!maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date)
                    || maps[field].Items.Any(key => key.Kind != PivotFieldValueKind.Date && key.Kind != PivotFieldValueKind.Blank))
                    throw new NotSupportedException("Relative-date filter materialization requires a date-valued pivot field.");
                QualifiedMaterializedPivotRelativeDateFilter(filter);
                if (!TryMaterializedPivotRelativeDateBounds(filter.Type!.Value, referenceDate,
                    out DateTime start, out DateTime end))
                    throw new NotSupportedException("The relative-date filter has no qualified calendar interval.");
                var included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key =>
                    key.Kind == PivotFieldValueKind.Date && key.Date!.Value >= start && key.Date.Value < end));
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(IsMaterializedPivotLabelComparison)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                bool numericItems = maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Number);
                if (maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date))
                    throw new NotSupportedException("Label-filter materialization has not qualified formatted date captions.");
                uint numberFormatId = 0;
                string? numberFormatCode = null;
                if (numericItems) {
                    var pivotFields = definition.PivotFields?.Elements<PivotField>().ToArray() ?? Array.Empty<PivotField>();
                    if (field >= pivotFields.Length)
                        throw new NotSupportedException("The numeric label field has no saved pivot format.");
                    numberFormatId = pivotFields[field].NumberFormatId?.Value ?? 0;
                    if (numberFormatId != 0) {
                        var workbookPart = _excelDocument.WorkbookPartRoot
                            ?? throw new InvalidOperationException("WorkbookPart is null");
                        if (!BuildNumberFormatCodeMap(workbookPart).TryGetValue(numberFormatId, out numberFormatCode)
                            || numberFormatCode is not ("#,##0" or "0.00" or "0.0" or "0.0%" or "0.000" or "$#,##0.00" or "\\$#,##0.00"))
                            throw new NotSupportedException("The numeric label format is not qualified for materialization.");
                    }
                }
                string DisplayCaption(PivotFieldValue key) {
                    if (captions.TryGetValue(field, out var names) && names.TryGetValue(key, out string? renamed))
                        return renamed;
                    return PivotMaterializedCaption(key, _excelDocument.DateSystem, numberFormatId, numberFormatCode);
                }
                string needle = filter.StringValue1?.Value
                    ?? throw new NotSupportedException("The label filter has no saved criterion.");
                var type = filter.Type!.Value;
                HashSet<PivotFieldValue> included;
                if (IsMaterializedPivotLabelRange(type)) {
                    if (numericItems)
                        throw new NotSupportedException("Numeric label ordering is not qualified for materialization.");
                    string? second = filter.StringValue2?.Value;
                    ValidateMaterializedPivotLabelRangePredicate(filter, needle, second);
                    if (!IsQualifiedPivotLabelOrderingText(needle)
                        || (second != null && !IsQualifiedPivotLabelOrderingText(second))
                        || maps[field].Items.Any(key => !IsQualifiedPivotLabelOrderingText(DisplayCaption(key))))
                        throw new NotSupportedException("Label range materialization requires ASCII alphabetic captions and criteria; localized text ordering is not qualified.");
                    CompareInfo compareInfo = CultureInfo.CurrentCulture.CompareInfo;
                    included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                        return MatchesMaterializedPivotLabelRange(DisplayCaption(key), needle, second, type, compareInfo);
                    }));
                } else {
                    string savedPattern = NormalizePivotFilterAutoFilterValue(type, needle);
                    var comparison = ResolveSingleFilterOperator(type);
                    ValidateMaterializedPivotLabelPredicate(filter, savedPattern, comparison);
                    var pattern = new Regex("\\A" + CreateFormulaWildcardPattern(savedPattern) + "\\z",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, FormulaRegexTimeout);
                    bool negate = comparison == FilterOperatorValues.NotEqual;
                    included = new HashSet<PivotFieldValue>(maps[field].Items.Where(key => {
                        return pattern.IsMatch(DisplayCaption(key)) != negate;
                    }));
                }
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            foreach (var filter in filters.Where(IsMaterializedPivotValueFilter)) {
                int field = QualifiedPivotFilterField(filter, cacheFields, axisFields);
                if (axisFields.Count != 1 || filter.MeasureField == null || filter.MeasureField.Value >= measures.Length)
                    throw new NotSupportedException("Value-filter materialization requires one ordinary axis field and a saved measure.");
                var type = filter.Type!.Value;
                double first = 0, second = 0;
                Top10? ranking = null;
                if (IsMaterializedPivotValueRanking(type)) {
                    ranking = QualifiedPivotRankingFilter(filter);
                } else if (type == PivotFilterValues.ValueBetween || type == PivotFilterValues.ValueNotBetween) {
                    (first, second) = QualifiedPivotValueRange(filter);
                } else {
                    var custom = QualifiedPivotCustomFilter(filter);
                    var comparison = ResolveSingleFilterOperator(type);
                    if ((custom.Operator?.Value ?? FilterOperatorValues.Equal) != comparison
                        || !TryFinitePivotThreshold(custom.Val?.Value, out first))
                        throw new NotSupportedException("The value filter does not have a qualified finite comparison threshold.");
                }
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
                var values = aggregates.Select(pair => (pair.Key, Value: pair.Value.GetValue(function).Value))
                    .Where(pair => pair.Value is double).Select(pair => (pair.Key, Value: (double)pair.Value!)).ToArray();
                HashSet<PivotFieldValue> included;
                if (ranking != null) {
                    if (aggregates.Count != 0 && values.Length == 0) {
                        var results = aggregates.Select(pair => (pair.Key, Result: pair.Value.GetValue(function))).ToArray();
                        if (results.Any(pair => pair.Result.Kind != ExcelCellDataKind.Error || pair.Result.Value is not string))
                            throw new NotSupportedException("Top/bottom ranking with no numeric or error aggregate is not qualified for materialization.");
                        if (type != PivotFilterValues.Count)
                            throw new NotSupportedException("Top/bottom percent and sum ranking over only error aggregates is not qualified for materialization.");
                        included = RankMaterializedPivotErrorValues(results, ranking);
                    } else {
                        included = RankMaterializedPivotValues(values, type, ranking);
                    }
                } else {
                    included = new HashSet<PivotFieldValue>(values.Where(pair =>
                        type == PivotFilterValues.ValueBetween ? pair.Value >= first && pair.Value <= second
                        : type == PivotFilterValues.ValueNotBetween ? pair.Value < first || pair.Value > second
                        : MatchesMaterializedPivotValueComparison(pair.Value, first, ResolveSingleFilterOperator(type)))
                        .Select(pair => pair.Key));
                }
                FilterMaterializedSourceRows(source, visibility.IncludedRows, included, field,
                    groupings, dateGroupings, manualGroupings, firstRow, lastRow, firstColumn, token);
            }

            if (filters.Any(filter => !IsMaterializedPivotFixedDateFilter(filter)
                && !IsMaterializedPivotCalendarFilter(filter)
                && !IsMaterializedPivotRelativeDateFilter(filter)
                && !IsMaterializedPivotLabelComparison(filter)
                && !IsMaterializedPivotValueFilter(filter)))
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

        private static bool IsMaterializedPivotFixedDateFilter(PivotFilter filter) {
            var type = filter.Type?.Value;
            return type == PivotFilterValues.DateEqual || type == PivotFilterValues.DateNotEqual
                || type == PivotFilterValues.DateNewerThan || type == PivotFilterValues.DateNewerThanOrEqual
                || type == PivotFilterValues.DateOlderThan || type == PivotFilterValues.DateOlderThanOrEqual
                || type == PivotFilterValues.DateBetween || type == PivotFilterValues.DateNotBetween;
        }

        private static bool QualifiedPivotWholeDay(PivotFilter filter) {
            var list = filter.GetFirstChild<PivotFilterExtensionList>();
            if (list == null) return false;
            var extensions = list.Elements<PivotFilterExtension>().ToArray();
            if (list.ChildElements.Count != 1 || extensions.Length != 1
                || extensions[0].Uri?.Value != WholeDayPivotFilterExtensionUri
                || extensions[0].ChildElements.Count != 1
                || extensions[0].GetFirstChild<X15.PivotFilter>() is not X15.PivotFilter setting
                || setting.UseWholeDay == null)
                throw new NotSupportedException("The fixed-date filter has an unsupported whole-day extension.");
            return setting.UseWholeDay.Value;
        }

        private static readonly PivotFilterValues[] MaterializedPivotCalendarMonthTypes = {
            PivotFilterValues.January, PivotFilterValues.February, PivotFilterValues.March,
            PivotFilterValues.April, PivotFilterValues.May, PivotFilterValues.June,
            PivotFilterValues.July, PivotFilterValues.August, PivotFilterValues.September,
            PivotFilterValues.October, PivotFilterValues.November, PivotFilterValues.December
        };

        private static readonly PivotFilterValues[] MaterializedPivotCalendarQuarterTypes = {
            PivotFilterValues.Quarter1, PivotFilterValues.Quarter2,
            PivotFilterValues.Quarter3, PivotFilterValues.Quarter4
        };

        private static bool IsMaterializedPivotCalendarFilter(PivotFilter filter)
            => filter.Type != null && TryMaterializedPivotCalendarMonths(filter.Type.Value, out _, out _);

        private static bool TryMaterializedPivotCalendarMonths(PivotFilterValues type, out int first, out int last) {
            int month = Array.FindIndex(MaterializedPivotCalendarMonthTypes, candidate => candidate == type);
            if (month >= 0) {
                first = last = month + 1;
                return true;
            }
            int quarter = Array.FindIndex(MaterializedPivotCalendarQuarterTypes, candidate => candidate == type);
            if (quarter >= 0) {
                first = quarter * 3 + 1;
                last = first + 2;
                return true;
            }
            first = last = 0;
            return false;
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
            PivotFilterValues type, CompareInfo compareInfo) {
            int firstComparison = compareInfo.Compare(label, first, CompareOptions.IgnoreCase);
            if (type == PivotFilterValues.CaptionGreaterThan) return firstComparison > 0;
            if (type == PivotFilterValues.CaptionGreaterThanOrEqual) return firstComparison >= 0;
            if (type == PivotFilterValues.CaptionLessThan) return firstComparison < 0;
            if (type == PivotFilterValues.CaptionLessThanOrEqual) return firstComparison <= 0;
            int secondComparison = compareInfo.Compare(label, second, CompareOptions.IgnoreCase);
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

        private static bool IsMaterializedPivotValueFilter(PivotFilter filter) {
            var type = filter.Type?.Value;
            return type == PivotFilterValues.ValueEqual || type == PivotFilterValues.ValueNotEqual
                || type == PivotFilterValues.ValueGreaterThan || type == PivotFilterValues.ValueGreaterThanOrEqual
                || type == PivotFilterValues.ValueLessThan || type == PivotFilterValues.ValueLessThanOrEqual
                || type == PivotFilterValues.ValueBetween || type == PivotFilterValues.ValueNotBetween
                || type == PivotFilterValues.Count || type == PivotFilterValues.Percent || type == PivotFilterValues.Sum;
        }

        private static bool IsMaterializedPivotValueRanking(PivotFilterValues type)
            => type == PivotFilterValues.Count || type == PivotFilterValues.Percent || type == PivotFilterValues.Sum;

        private static bool TryFinitePivotThreshold(string? text, out double value)
            => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.IsNaN(value) && !double.IsInfinity(value);

        private static (double First, double Second) QualifiedPivotValueRange(PivotFilter filter) {
            PivotFilterValues type = filter.Type!.Value;
            TryResolveBetweenFilter(type, out var firstOperator, out var secondOperator, out bool matchAll);
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (columns.Length != 1 || columns[0].ColumnId?.Value != 0 || columns[0].ChildElements.Count != 1
                || columns[0].GetFirstChild<CustomFilters>() is not CustomFilters custom
                || (custom.And?.Value == true) != matchAll || custom.ChildElements.Count != 2)
                throw new NotSupportedException("The value range filter has no qualified saved predicates.");
            var predicates = custom.Elements<CustomFilter>().ToArray();
            if (predicates.Length != 2 || (predicates[0].Operator?.Value ?? FilterOperatorValues.Equal) != firstOperator
                || (predicates[1].Operator?.Value ?? FilterOperatorValues.Equal) != secondOperator
                || !TryFinitePivotThreshold(predicates[0].Val?.Value, out double first)
                || !TryFinitePivotThreshold(predicates[1].Val?.Value, out double second)
                || first > second
                || (filter.StringValue1 != null && (!TryFinitePivotThreshold(filter.StringValue1.Value, out double savedFirst) || savedFirst != first))
                || (filter.StringValue2 != null && (!TryFinitePivotThreshold(filter.StringValue2.Value, out double savedSecond) || savedSecond != second)))
                throw new NotSupportedException("The value range filter has no qualified finite bounds.");
            return (first, second);
        }

        private static Top10 QualifiedPivotRankingFilter(PivotFilter filter) {
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (columns.Length != 1 || columns[0].ColumnId?.Value != 0 || columns[0].ChildElements.Count != 1
                || columns[0].GetFirstChild<Top10>() is not Top10 ranking
                || !TryFinitePivotThreshold(ranking.Val?.Value.ToString(CultureInfo.InvariantCulture), out double threshold)
                || threshold <= 0 || (filter.Type!.Value == PivotFilterValues.Count && (threshold != Math.Truncate(threshold) || threshold > 100_000))
                || (filter.Type.Value == PivotFilterValues.Percent && (threshold > 100 || ranking.Percent?.Value != true))
                || (filter.Type.Value != PivotFilterValues.Percent && ranking.Percent?.Value == true)
                || (filter.StringValue1 != null && (!TryFinitePivotThreshold(filter.StringValue1.Value, out double saved) || saved != threshold)))
                throw new NotSupportedException("The top/bottom value filter has no qualified saved ranking rule.");
            return ranking;
        }

        private static HashSet<PivotFieldValue> RankMaterializedPivotValues(
            (PivotFieldValue Key, double Value)[] values, PivotFilterValues type, Top10 ranking) {
            if (values.Length == 0) return new HashSet<PivotFieldValue>();
            bool top = ranking.Top?.Value != false;
            Array.Sort(values, (left, right) => top
                ? right.Value.CompareTo(left.Value) : left.Value.CompareTo(right.Value));
            double requested = ranking.Val!.Value;
            double target = type == PivotFilterValues.Percent
                ? values.Sum(pair => pair.Value) * (requested / 100d) : requested;
            if (double.IsNaN(target) || double.IsInfinity(target))
                throw new NotSupportedException("The top/bottom percent total exceeds the qualified numeric range.");
            if (type == PivotFilterValues.Percent && target == 0)
                return new HashSet<PivotFieldValue>(values.Select(pair => pair.Key));
            int cutoffIndex = 0;
            if (type == PivotFilterValues.Count) {
                cutoffIndex = Math.Min(values.Length, (int)requested) - 1;
            } else {
                double accumulated = 0;
                for (; cutoffIndex < values.Length - 1; cutoffIndex++) {
                    accumulated += values[cutoffIndex].Value;
                    if (double.IsNaN(accumulated) || double.IsInfinity(accumulated))
                        throw new NotSupportedException("The top/bottom ranking total exceeds the qualified numeric range.");
                    if (target < 0 ? accumulated <= target : accumulated >= target) break;
                }
            }
            double cutoff = values[cutoffIndex].Value;
            return new HashSet<PivotFieldValue>(values.Where(pair => top
                ? pair.Value >= cutoff : pair.Value <= cutoff).Select(pair => pair.Key));
        }

        private static HashSet<PivotFieldValue> RankMaterializedPivotErrorValues(
            (PivotFieldValue Key, ExcelCellData Result)[] results, Top10 ranking) {
            var values = new (PivotFieldValue Key, byte Code)[results.Length];
            for (int index = 0; index < results.Length; index++) {
                string error = (string)results[index].Result.Value!;
                if (!ExcelErrorCode.TryGetCode(error, out byte code) || code > 0x2a)
                    throw new NotSupportedException("The pivot error ranking contains an unqualified error value.");
                values[index] = (results[index].Key, code);
            }
            bool top = ranking.Top?.Value != false;
            Array.Sort(values, (left, right) => top
                ? left.Code.CompareTo(right.Code) : right.Code.CompareTo(left.Code));
            byte cutoff = values[Math.Min(values.Length, (int)ranking.Val!.Value) - 1].Code;
            return new HashSet<PivotFieldValue>(values.Where(pair => top
                ? pair.Code <= cutoff : pair.Code >= cutoff).Select(pair => pair.Key));
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
