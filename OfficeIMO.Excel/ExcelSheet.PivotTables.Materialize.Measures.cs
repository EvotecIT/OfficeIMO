using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct PivotMaterializedAxisEntry {
            internal PivotMaterializedAxisEntry(int key, int measure) { Key = key; Measure = measure; }
            internal int Key { get; }
            internal int Measure { get; }
        }

        private static List<PivotMaterializedAxisEntry> BuildMaterializedAxisEntries(PivotMaterializationAxis axis, int keys, int measures, bool total) {
            int axisMeasures = axis.HasValues ? measures : 1;
            var result = new List<PivotMaterializedAxisEntry>((keys + (total ? 1 : 0)) * axisMeasures);
            if (axis.HasValues && axis.Fields[0] == -2) {
                for (int measure = 0; measure < axisMeasures; measure++)
                    for (int key = 0; key < keys; key++) result.Add(new PivotMaterializedAxisEntry(key, measure));
            } else {
                for (int key = 0; key < keys; key++)
                    for (int measure = 0; measure < axisMeasures; measure++) result.Add(new PivotMaterializedAxisEntry(key, measure));
            }
            if (total) for (int measure = 0; measure < axisMeasures; measure++) result.Add(new PivotMaterializedAxisEntry(-1, measure));
            return result;
        }

        private void FillMaterializedPivotMeasures(PivotMaterializationPlan plan, ExcelSheet source, int firstRow, int firstColumn, int lastRow,
            IReadOnlyList<PivotFieldValues> maps, PivotMaterializationAxis rowAxis, PivotMaterializationAxis columnAxis,
            DataField[] measures, bool rowTotal, bool columnTotal, int dataRow, int dataColumn, CancellationToken token) {
            int rowField = rowAxis.RealField, columnField = columnAxis.RealField;
            int rows = rowField < 0 ? 1 : maps[rowField].Items.Count;
            int columns = columnField < 0 ? 1 : maps[columnField].Items.Count;
            var rowIndices = rowField < 0 ? null : IndexPivotMaterializationKeys(maps[rowField]);
            var columnIndices = columnField < 0 ? null : IndexPivotMaterializationKeys(maps[columnField]);
            var groups = new Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator[]>();
            var rowTotals = Enumerable.Range(0, rows).Select(_ => CreatePivotMeasureAccumulators(measures.Length)).ToArray();
            var columnTotals = Enumerable.Range(0, columns).Select(_ => CreatePivotMeasureAccumulators(measures.Length)).ToArray();
            var grandTotal = CreatePivotMeasureAccumulators(measures.Length);
            var functions = measures.Select(m => (m.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum).ToOfficeEnum()).ToArray();
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                int rowKey = rowIndices == null ? 0 : rowIndices[source.GetPivotFieldValue(row, firstColumn + rowField, null)];
                int columnKey = columnIndices == null ? 0 : columnIndices[source.GetPivotFieldValue(row, firstColumn + columnField, null)];
                if (!groups.TryGetValue((rowKey, columnKey), out var group)) {
                    if (groups.Count >= 100_000 / measures.Length)
                        throw new InvalidOperationException("The pivot exceeds 100,000 observed aggregate states across its measures.");
                    groups.Add((rowKey, columnKey), group = CreatePivotMeasureAccumulators(measures.Length));
                }
                for (int measure = 0; measure < measures.Length; measure++) {
                    var cell = source.TryGetExistingCell(row, firstColumn + (int)measures[measure].Field!.Value);
                    var value = source.GetCellValueSnapshot(cell);
                    bool error = cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
                    group[measure].Add(value.Value, error);
                    rowTotals[rowKey][measure].Add(value.Value, error);
                    columnTotals[columnKey][measure].Add(value.Value, error);
                    grandTotal[measure].Add(value.Value, error);
                }
            }
            var rowEntries = BuildMaterializedAxisEntries(rowAxis, rows, measures.Length, rowTotal);
            var columnEntries = BuildMaterializedAxisEntries(columnAxis, columns, measures.Length, columnTotal);
            var values = new ExcelCellData?[plan.Bottom - plan.Top + 1, plan.Right - plan.Left + 1];
            var fields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            string totalCaption = plan.Definition.GrandTotalCaption?.Value ?? "Grand Total";
            string valuesCaption = plan.Definition.DataCaption?.Value ?? "Values";
            for (int level = 0; level < rowAxis.Fields.Length; level++) {
                int field = rowAxis.Fields[level];
                values[dataRow - 1, level] = PivotMaterializedText(field == -2 ? valuesCaption : fields[field].Name?.Value ?? "");
                for (int index = 0; index < rowEntries.Count; index++) {
                    var entry = rowEntries[index];
                    values[dataRow + index, level] = field == -2 ? PivotMaterializedText(measures[entry.Measure].Name?.Value ?? "")
                        : entry.Key < 0 ? PivotMaterializedText(totalCaption) : PivotMaterializedKey(maps[field].Items[entry.Key]);
                }
            }
            if (columnField >= 0) values[0, dataColumn] = PivotMaterializedText(plan.Definition.ColumnHeaderCaption?.Value ?? fields[columnField].Name?.Value ?? "");
            for (int level = 0; level < columnAxis.Fields.Length; level++) {
                int field = columnAxis.Fields[level];
                int headerRow = columnField >= 0 ? level + 1 : level;
                for (int index = 0; index < columnEntries.Count; index++) {
                    var entry = columnEntries[index];
                    values[headerRow, dataColumn + index] = field == -2 ? PivotMaterializedText(measures[entry.Measure].Name?.Value ?? "")
                        : entry.Key < 0 ? PivotMaterializedText(totalCaption) : PivotMaterializedKey(maps[field].Items[entry.Key]);
                }
            }
            if (columnAxis.Fields.Length == 0) values[dataRow - 1, dataColumn] = PivotMaterializedText(valuesCaption);
            for (int row = 0; row < rowEntries.Count; row++) {
                token.ThrowIfCancellationRequested();
                var rowEntry = rowEntries[row];
                for (int column = 0; column < columnEntries.Count; column++) {
                    var columnEntry = columnEntries[column];
                    int measure = rowAxis.HasValues ? rowEntry.Measure : columnEntry.Measure;
                    ExcelPivotAggregateAccumulator[]? aggregate = rowEntry.Key < 0
                        ? columnEntry.Key < 0 ? grandTotal : columnTotals[columnEntry.Key]
                        : columnEntry.Key < 0 ? rowTotals[rowEntry.Key]
                        : groups.TryGetValue((rowEntry.Key, columnEntry.Key), out var group) ? group : null;
                    values[dataRow + row, dataColumn + column] = aggregate == null
                        ? new ExcelCellData(ExcelCellDataKind.Number, 0d) : aggregate[measure].GetValue(functions[measure]);
                }
            }
            plan.Values = values;
            var definition = plan.Definition;
            definition.Location = new Location { Reference = $"{A1.CellReference(plan.Top, plan.Left)}:{A1.CellReference(plan.Bottom, plan.Right)}",
                FirstHeaderRow = columnField >= 0 || !columnAxis.HasValues ? 1U : 0U, FirstDataRow = (uint)dataRow, FirstDataColumn = (uint)dataColumn };
            definition.Compact = false;
            definition.CompactData = false;
            definition.OutlineData = false;
            definition.DataOnRows = rowAxis.HasValues;
            NormalizeMaterializedPivotFields(definition, maps, rowField, columnField);
            definition.RowItems = new RowItems { Count = (uint)rowEntries.Count };
            definition.ColumnItems = new ColumnItems { Count = (uint)columnEntries.Count };
            foreach (var entry in rowEntries) definition.RowItems.AppendChild(CreateMaterializedMeasureAxisItem(rowAxis, entry));
            foreach (var entry in columnEntries) definition.ColumnItems.AppendChild(CreateMaterializedMeasureAxisItem(columnAxis, entry));
        }

        private static ExcelPivotAggregateAccumulator[] CreatePivotMeasureAccumulators(int count)
            => Enumerable.Range(0, count).Select(_ => new ExcelPivotAggregateAccumulator()).ToArray();

        private static RowItem CreateMaterializedMeasureAxisItem(PivotMaterializationAxis axis, PivotMaterializedAxisEntry entry) {
            var item = new RowItem();
            if (axis.HasValues) item.Index = (uint)entry.Measure;
            if (entry.Key < 0) {
                item.ItemType = ItemValues.Grand;
                item.AppendChild(new MemberPropertyIndex { Val = axis.HasValues ? entry.Measure : 0 });
            } else {
                foreach (int field in axis.Fields) item.AppendChild(new MemberPropertyIndex { Val = field == -2 ? entry.Measure : entry.Key });
            }
            return item;
        }
    }
}
