using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotValueFilterGroup {
            internal Dictionary<PivotFieldValue, PivotValueFilterGroup>? Children;
            internal ExcelPivotAggregateAccumulator? Aggregate;
            internal bool Included;
        }

        private void ApplyMaterializedPivotValueFilter(
            ExcelSheet source, bool[] includedRows, int[] fieldPrefix, DataField measure,
            PivotFilterValues type, double first, double second, Top10? ranking,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings,
            IReadOnlyDictionary<int, PivotManualGrouping> manualGroupings,
            int firstRow, int lastRow, int firstColumn, int limit, CancellationToken token) {
            if (fieldPrefix.Length == 0)
                throw new NotSupportedException("The value filter requires a row or column axis field.");

            var root = new PivotValueFilterGroup();
            int groupCount = 0;
            int measureColumn = firstColumn + (int)measure.Field!.Value;
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                if (!includedRows[row - firstRow - 1]) continue;
                var group = root;
                foreach (int field in fieldPrefix) {
                    var key = MaterializedPivotAxisKey(source, row, firstColumn + field, field,
                        groupings, dateGroupings, manualGroupings);
                    var children = group.Children ??= new Dictionary<PivotFieldValue, PivotValueFilterGroup>();
                    if (!children.TryGetValue(key, out var child)) {
                        if (groupCount >= 100_000 || groupCount >= limit)
                            throw new InvalidOperationException("The value filter exceeds the pivot aggregate budget.");
                        child = new PivotValueFilterGroup();
                        children.Add(key, child);
                        groupCount++;
                    }
                    group = child;
                }
                if (group.Aggregate == null) group.Aggregate = new ExcelPivotAggregateAccumulator();
                var cell = source.TryGetExistingCell(row, measureColumn);
                group.Aggregate.Add(source.GetCellValueSnapshot(cell).Value,
                    cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error);
            }
            if (groupCount == 0) return;

            // A value filter ranks or compares children within each parent on its own axis.
            // The outermost field has one implicit root parent.
            var parents = new List<PivotValueFilterGroup> { root };
            for (int depth = 1; depth < fieldPrefix.Length; depth++) {
                var next = new List<PivotValueFilterGroup>();
                foreach (var parent in parents) next.AddRange(parent.Children!.Values);
                parents = next;
            }
            var function = (measure.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum).ToOfficeEnum();
            foreach (var parent in parents) {
                token.ThrowIfCancellationRequested();
                var aggregates = parent.Children!;
                var values = aggregates.Select(pair => (pair.Key, Value: pair.Value.Aggregate!.GetValue(function).Value))
                    .Where(pair => pair.Value is double).Select(pair => (pair.Key, Value: (double)pair.Value!)).ToArray();
                HashSet<PivotFieldValue> included;
                if (ranking != null) {
                    if (aggregates.Count != 0 && values.Length == 0) {
                        var results = aggregates.Select(pair => (pair.Key, Result: pair.Value.Aggregate!.GetValue(function))).ToArray();
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
                foreach (var pair in aggregates) pair.Value.Included = included.Contains(pair.Key);
            }

            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                int index = row - firstRow - 1;
                if (!includedRows[index]) continue;
                var group = root;
                foreach (int field in fieldPrefix) {
                    var key = MaterializedPivotAxisKey(source, row, firstColumn + field, field,
                        groupings, dateGroupings, manualGroupings);
                    if (group.Children == null || !group.Children.TryGetValue(key, out var child))
                        throw new InvalidOperationException("The value-filter source changed during materialization.");
                    group = child;
                }
                if (!group.Included) includedRows[index] = false;
            }
        }
    }
}
