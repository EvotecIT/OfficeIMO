using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void FillMaterializedHierarchy(PivotMaterializationPlan plan, IReadOnlyList<PivotFieldValues> maps,
            PivotHierarchyAxis rows, PivotHierarchyAxis columns, DataField[] measures, int dataRow, int dataColumn,
            Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator[]> groups, CancellationToken token) {
            var definition = plan.Definition;
            var fields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            var values = new ExcelCellData?[plan.Bottom - plan.Top + 1, plan.Right - plan.Left + 1];
            string caption = measures.Length == 1 ? measures[0].Name?.Value ?? "Values" : definition.DataCaption?.Value ?? "Values";
            string totalCaption = definition.GrandTotalCaption?.Value ?? "Grand Total";
            for (int level = 0; level < rows.Layout.Fields.Length; level++) {
                int field = rows.Layout.Fields[level];
                values[dataRow - 1, level] = PivotMaterializedText(field == -2 ? definition.DataCaption?.Value ?? "Values" : fields[field].Name?.Value ?? "");
            }
            if (columns.Layout.RealFields.Length > 0) {
                values[0, dataColumn] = PivotMaterializedText(definition.ColumnHeaderCaption?.Value ?? fields[columns.Layout.RealFields[0]].Name?.Value ?? "");
                if (measures.Length == 1) {
                    if (rows.Layout.RealFields.Length > 0) values[0, 0] = PivotMaterializedText(caption);
                    else values[dataRow, 0] = PivotMaterializedText(caption);
                }
            } else if (columns.Layout.Fields.Length == 0) values[dataRow - 1, dataColumn] = PivotMaterializedText(caption);
            void Labels(PivotHierarchyAxis axis, PivotHierarchyEntry entry, Action<int, ExcelCellData> put) {
                int[] keys = MaterializedHierarchyKeys(entry.Node);
                int depth = 0;
                bool grandLabel = false;
                for (int level = 0; level < axis.Layout.Fields.Length; level++) {
                    int field = axis.Layout.Fields[level];
                    if (field == -2) { put(level, PivotMaterializedText(measures[entry.Measure].Name?.Value ?? "")); continue; }
                    if (entry.Type == ItemValues.Grand) {
                        if (!grandLabel) { put(level, PivotMaterializedText(totalCaption)); grandLabel = true; }
                    } else if (depth < keys.Length) {
                        var key = maps[field].Items[keys[depth++]];
                        var label = PivotMaterializedKey(key);
                        if (entry.Type == ItemValues.Default && depth == keys.Length) {
                            string text = key.Kind == PivotFieldValueKind.Blank ? "(blank)"
                                : key.Kind == PivotFieldValueKind.Boolean ? key.Boolean == true ? "TRUE" : "FALSE" : key.Text;
                            label = PivotMaterializedText(text + " Total");
                        }
                        put(level, label);
                    }
                }
            }
            for (int row = 0; row < rows.Entries.Count; row++) {
                token.ThrowIfCancellationRequested();
                int position = dataRow + row;
                Labels(rows, rows.Entries[row], (level, value) => values[position, level] = value);
            }
            for (int column = 0; column < columns.Entries.Count; column++) {
                int position = dataColumn + column;
                int offset = columns.Layout.RealFields.Length > 0 ? 1 : 0;
                Labels(columns, columns.Entries[column], (level, value) => values[level + offset, position] = value);
            }
            var functions = measures.Select(measure => (measure.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum).ToOfficeEnum()).ToArray();
            for (int row = 0; row < rows.Entries.Count; row++) {
                token.ThrowIfCancellationRequested();
                var rowEntry = rows.Entries[row];
                for (int column = 0; column < columns.Entries.Count; column++) {
                    var columnEntry = columns.Entries[column];
                    int measure = rows.Layout.HasValues ? rowEntry.Measure : columnEntry.Measure;
                    values[dataRow + row, dataColumn + column] = groups.TryGetValue((rowEntry.Node.Id, columnEntry.Node.Id), out var aggregate)
                        ? aggregate[measure].GetValue(functions[measure]) : new ExcelCellData(ExcelCellDataKind.Number, 0d);
                }
            }
            plan.Values = values;
            definition.Location = new Location { Reference = $"{A1.CellReference(plan.Top, plan.Left)}:{A1.CellReference(plan.Bottom, plan.Right)}",
                FirstHeaderRow = columns.Layout.RealFields.Length > 0 || !columns.Layout.HasValues ? 1U : 0U,
                FirstDataRow = (uint)dataRow, FirstDataColumn = (uint)dataColumn };
            definition.Compact = false;
            definition.CompactData = false;
            definition.OutlineData = false;
            definition.DataOnRows = rows.Layout.HasValues;
            NormalizeMaterializedPivotFields(definition, maps, rows);
            NormalizeMaterializedPivotFields(definition, maps, columns);
            definition.RowItems = new RowItems { Count = (uint)rows.Entries.Count };
            definition.ColumnItems = new ColumnItems { Count = (uint)columns.Entries.Count };
            foreach (var entry in rows.Entries) definition.RowItems.AppendChild(CreateMaterializedHierarchyItem(rows, entry));
            foreach (var entry in columns.Entries) definition.ColumnItems.AppendChild(CreateMaterializedHierarchyItem(columns, entry));
        }

        private static void NormalizeMaterializedPivotFields(PivotTableDefinition definition, IReadOnlyList<PivotFieldValues> maps, PivotHierarchyAxis axis) {
            var fields = definition.PivotFields!.Elements<PivotField>().ToArray();
            for (int depth = 0; depth < axis.Layout.RealFields.Length; depth++) {
                int field = axis.Layout.RealFields[depth];
                bool subtotal = axis.Subtotals[depth];
                var items = new Items { Count = (uint)(maps[field].Items.Count + (subtotal ? 1 : 0)) };
                for (int index = 0; index < maps[field].Items.Count; index++) items.AppendChild(new Item { Index = (uint)index });
                if (subtotal) items.AppendChild(new Item { ItemType = ItemValues.Default });
                fields[field].Items = items;
                fields[field].DefaultSubtotal = subtotal;
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
                fields[field].SubtotalTop = false;
                fields[field].SortType = FieldSortValues.Manual;
            }
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
