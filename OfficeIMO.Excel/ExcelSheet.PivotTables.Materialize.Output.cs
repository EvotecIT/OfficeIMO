using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void FillMaterializedPivot(PivotMaterializationPlan plan, ExcelSheet source, int firstRow, int firstColumn, int lastRow,
            IReadOnlyList<PivotFieldValues> maps, int rowField, int columnField, int measureField, ExcelPivotDataFunction function,
            bool rowTotal, bool columnTotal, int dataRow, int dataColumn, CancellationToken token) {
            int rows = rowField < 0 ? 1 : maps[rowField].Items.Count;
            int columns = columnField < 0 ? 1 : maps[columnField].Items.Count;
            var rowIndices = rowField < 0 ? null : IndexPivotMaterializationKeys(maps[rowField]);
            var columnIndices = columnField < 0 ? null : IndexPivotMaterializationKeys(maps[columnField]);
            var groups = new Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator>();
            var rowTotals = Enumerable.Range(0, rows).Select(_ => new ExcelPivotAggregateAccumulator()).ToArray();
            var columnTotals = Enumerable.Range(0, columns).Select(_ => new ExcelPivotAggregateAccumulator()).ToArray();
            var grandTotal = new ExcelPivotAggregateAccumulator();
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                int rowKey = rowIndices == null ? 0 : rowIndices[source.GetPivotFieldValue(row, firstColumn + rowField, null)];
                int columnKey = columnIndices == null ? 0 : columnIndices[source.GetPivotFieldValue(row, firstColumn + columnField, null)];
                if (!groups.TryGetValue((rowKey, columnKey), out var group)) {
                    if (groups.Count >= 100_000) throw new InvalidOperationException("The pivot exceeds 100,000 observed aggregate groups.");
                    groups.Add((rowKey, columnKey), group = new ExcelPivotAggregateAccumulator());
                }
                var cell = source.TryGetExistingCell(row, firstColumn + measureField);
                var value = source.GetCellValueSnapshot(cell);
                bool error = cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
                group.Add(value.Value, error);
                rowTotals[rowKey].Add(value.Value, error);
                columnTotals[columnKey].Add(value.Value, error);
                grandTotal.Add(value.Value, error);
            }
            var definition = plan.Definition;
            string caption = definition.DataFields!.Elements<DataField>().Single().Name?.Value ?? "Values";
            string totalCaption = definition.GrandTotalCaption?.Value ?? "Grand Total";
            var values = new ExcelCellData?[plan.Bottom - plan.Top + 1, plan.Right - plan.Left + 1];
            if (columnField >= 0) {
                if (rowField >= 0) values[0, 0] = PivotMaterializedText(caption);
                values[0, dataColumn] = PivotMaterializedText(definition.ColumnHeaderCaption?.Value
                    ?? plan.Cache.CacheFields!.Elements<CacheField>().ElementAt(columnField).Name?.Value ?? "");
                if (rowField < 0) values[dataRow, 0] = PivotMaterializedText(caption);
                for (int column = 0; column < columns; column++) values[1, dataColumn + column] = PivotMaterializedKey(maps[columnField].Items[column]);
                if (columnTotal) values[1, dataColumn + columns] = PivotMaterializedText(totalCaption);
            } else values[0, dataColumn] = PivotMaterializedText(caption);
            if (rowField >= 0) {
                values[dataRow - 1, 0] = PivotMaterializedText(plan.Cache.CacheFields!.Elements<CacheField>().ElementAt(rowField).Name?.Value ?? "");
                for (int row = 0; row < rows; row++) values[dataRow + row, 0] = PivotMaterializedKey(maps[rowField].Items[row]);
                if (rowTotal) values[dataRow + rows, 0] = PivotMaterializedText(totalCaption);
            }
            for (int row = 0; row < rows; row++) {
                token.ThrowIfCancellationRequested();
                for (int column = 0; column < columns; column++) {
                    values[dataRow + row, dataColumn + column] = groups.TryGetValue((row, column), out var group)
                        ? group.GetValue(function) : new ExcelCellData(ExcelCellDataKind.Number, 0d);
                }
                if (columnTotal) values[dataRow + row, dataColumn + columns] = rowTotals[row].GetValue(function);
            }
            if (rowTotal) {
                for (int column = 0; column < columns; column++) values[dataRow + rows, dataColumn + column] = columnTotals[column].GetValue(function);
                if (columnTotal) values[dataRow + rows, dataColumn + columns] = grandTotal.GetValue(function);
            }
            plan.Values = values;
            definition.Location = new Location { Reference = $"{A1.CellReference(plan.Top, plan.Left)}:{A1.CellReference(plan.Bottom, plan.Right)}",
                FirstHeaderRow = 1U, FirstDataRow = (uint)dataRow, FirstDataColumn = (uint)dataColumn };
            definition.CompactData = false;
            definition.OutlineData = false;
            definition.DataOnRows = false;
            NormalizeMaterializedPivotFields(definition, maps, rowField, columnField);
            definition.RowItems = new RowItems { Count = (uint)(rows + (rowTotal ? 1 : 0)) };
            definition.ColumnItems = new ColumnItems { Count = (uint)(columns + (columnTotal ? 1 : 0)) };
            for (int row = 0; row < rows; row++) definition.RowItems.AppendChild(CreateMaterializedPivotAxisItem(row, hasField: rowField >= 0));
            if (rowTotal) definition.RowItems.AppendChild(CreateMaterializedPivotAxisItem(0, true));
            for (int column = 0; column < columns; column++) definition.ColumnItems.AppendChild(CreateMaterializedPivotAxisItem(column, hasField: columnField >= 0));
            if (columnTotal) definition.ColumnItems.AppendChild(CreateMaterializedPivotAxisItem(0, true));
        }

        private static void NormalizeMaterializedPivotFields(PivotTableDefinition definition, IReadOnlyList<PivotFieldValues> maps, int rowField, int columnField) {
            var fields = definition.PivotFields!.Elements<PivotField>().ToArray();
            foreach (int field in new[] { rowField, columnField }.Where(f => f >= 0)) {
                var items = new Items { Count = (uint)(maps[field].Items.Count + 1) };
                for (int index = 0; index < maps[field].Items.Count; index++) items.AppendChild(new Item { Index = (uint)index });
                items.AppendChild(new Item { ItemType = ItemValues.Default });
                fields[field].Items = items;
                // The saved default item requires the matching subtotal setting.
                // With one field per axis there are no intermediate subtotal rows.
                fields[field].DefaultSubtotal = true;
                fields[field].SumSubtotal = false;
                fields[field].CountASubtotal = false;
                fields[field].AverageSubTotal = false;
                fields[field].MaxSubtotal = false;
                fields[field].MinSubtotal = false;
                fields[field].ApplyProductInSubtotal = false;
                fields[field].CountSubtotal = false;
                fields[field].ApplyStandardDeviationInSubtotal = false;
                fields[field].ApplyStandardDeviationPInSubtotal = false;
                fields[field].ApplyVarianceInSubtotal = false;
                fields[field].ApplyVariancePInSubtotal = false;
                fields[field].Compact = false;
                fields[field].Outline = false;
                fields[field].SortType = FieldSortValues.Manual;
            }
        }

        private static RowItem CreateMaterializedPivotAxisItem(int index, bool grand = false, bool hasField = true) {
            var item = new RowItem();
            if (hasField) item.AppendChild(new MemberPropertyIndex { Val = index });
            if (grand) item.ItemType = ItemValues.Grand;
            return item;
        }

        private static Dictionary<PivotFieldValue, int> IndexPivotMaterializationKeys(PivotFieldValues values) {
            var result = new Dictionary<PivotFieldValue, int>(values.Items.Count);
            for (int index = 0; index < values.Items.Count; index++) result.Add(values.Items[index], index);
            return result;
        }

        private static ExcelCellData PivotMaterializedText(string text) => new(ExcelCellDataKind.Text, text);
        private static ExcelCellData PivotMaterializedKey(PivotFieldValue key) => key.Kind switch {
            PivotFieldValueKind.Blank => PivotMaterializedText("(blank)"),
            PivotFieldValueKind.Boolean => new(ExcelCellDataKind.Boolean, key.Boolean),
            PivotFieldValueKind.Number => new(ExcelCellDataKind.Number, key.Number),
            _ => PivotMaterializedText(key.Text)
        };

        private void ValidatePivotMaterializationDestination(PivotTablePart part, ExcelSheet source, int r1, int c1, int r2, int c2,
            int top, int left, int bottom, int right, int oldBottom, int oldRight, CancellationToken token) {
            bool Intersects(string? range) => range != null && A1.TryParseRange(range, out int rr1, out int cc1, out int rr2, out int cc2)
                && rr1 <= bottom && rr2 >= top && cc1 <= right && cc2 >= left;
            if (ReferenceEquals(_worksheetPart, source._worksheetPart) && r1 <= bottom && r2 >= top && c1 <= right && c2 >= left)
                throw new InvalidOperationException("The pivot output overlaps its source range.");
            if (_worksheetPart.PivotTableParts.Any(p => !ReferenceEquals(p, part) && Intersects(p.PivotTableDefinition?.Location?.Reference?.Value))
                || _worksheetPart.TableDefinitionParts.Any(p => Intersects(p.Table?.Reference?.Value))
                || WorksheetRoot.GetFirstChild<MergeCells>()?.Elements<MergeCell>().Any(m => Intersects(m.Reference?.Value)) == true)
                throw new InvalidOperationException("The pivot output overlaps another pivot, table or merged range.");
            for (int row = top; row <= bottom; row++) {
                token.ThrowIfCancellationRequested();
                for (int column = left; column <= right; column++) {
                    var cell = TryGetExistingCell(row, column);
                    if (cell?.CellFormula != null || ((row > oldBottom || column > oldRight) && GetCellValueSnapshot(cell).Kind != ExcelCellDataKind.Blank))
                        throw new InvalidOperationException($"Pivot output would overwrite unrelated content at {A1.CellReference(row, column)}.");
                }
            }
        }
    }
}
