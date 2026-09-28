using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void FillMaterializedHierarchy(PivotMaterializationPlan plan, IReadOnlyList<PivotFieldValues> maps,
            PivotHierarchyAxis rows, PivotHierarchyAxis columns, PivotMaterializationVisibility visibility,
            DataField[] measures, int dataRow, int dataColumn,
            Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator[]> groups,
            bool dateHierarchy, bool dateColumnHierarchy, CancellationToken token) {
            var definition = plan.Definition;
            var fields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            var values = new ExcelCellData?[plan.Bottom - plan.Top + 1, plan.Right - plan.Left + 1];
            bool dateRowValues = dateHierarchy && rows.Layout.HasValues;
            int rowValuesPosition = Array.IndexOf(rows.Layout.Fields, -2);
            string caption = measures.Length == 1 ? measures[0].Name?.Value ?? "Values" : definition.DataCaption?.Value ?? "Values";
            string totalCaption = definition.GrandTotalCaption?.Value ?? "Grand Total";
            for (int level = 0; level < rows.Layout.Fields.Length; level++) {
                int field = rows.Layout.Fields[level];
                values[dataRow - 1, level] = PivotMaterializedText(field == -2 ? definition.DataCaption?.Value ?? "Values" : fields[field].Name?.Value ?? "");
            }
            if (columns.Layout.RealFields.Length > 0) {
                values[0, dataColumn] = PivotMaterializedText(definition.ColumnHeaderCaption?.Value ?? fields[columns.Layout.RealFields[0]].Name?.Value ?? "");
                if (dateColumnHierarchy) {
                    for (int level = 1; level < columns.Layout.RealFields.Length; level++)
                        values[0, dataColumn + level] = PivotMaterializedText(fields[columns.Layout.RealFields[level]].Name?.Value ?? "");
                }
                if (measures.Length == 1) {
                    if (rows.Layout.RealFields.Length > 0) values[0, 0] = PivotMaterializedText(caption);
                    else values[dataRow, 0] = PivotMaterializedText(caption);
                }
            } else if (columns.Layout.Fields.Length == 0 && !dateRowValues) values[dataRow - 1, dataColumn] = PivotMaterializedText(caption);
            void Labels(PivotHierarchyAxis axis, PivotHierarchyEntry entry, int[]? rowKeys, bool dateRowLayout, bool firstMeasureRow,
                Action<int, ExcelCellData, bool> put) {
                if (dateRowLayout && axis.Layout.HasValues && entry.Type == ItemValues.Grand) {
                    put(0, PivotMaterializedText("Total " + (measures[entry.Measure].Name?.Value ?? "")), false);
                    return;
                }
                int[] keys = rowKeys ?? MaterializedHierarchyKeys(entry.Node);
                int depth = 0;
                bool grandLabel = false;
                for (int level = 0; level < axis.Layout.Fields.Length; level++) {
                    int field = axis.Layout.Fields[level];
                    if (field == -2) {
                        string measureCaption = measures[entry.Measure].Name?.Value ?? "";
                        if (!dateRowLayout || entry.Type == ItemValues.Data && firstMeasureRow)
                            put(level, PivotMaterializedText(measureCaption), false);
                        continue;
                    }
                    if (entry.Type == ItemValues.Grand) {
                        if (!grandLabel) { put(level, PivotMaterializedText(totalCaption), false); grandLabel = true; }
                    } else if (depth < keys.Length) {
                        var key = maps[field].Items[keys[depth++]];
                        var label = PivotMaterializedKey(key, plan.SourceDateSystem);
                        bool subtotalLabel = entry.Type == ItemValues.Default && depth == keys.Length;
                        if (subtotalLabel) {
                            string text = key.Kind == PivotFieldValueKind.Blank ? "(blank)"
                                : key.Kind == PivotFieldValueKind.Boolean ? key.Boolean == true ? "TRUE" : "FALSE"
                                : key.Kind == PivotFieldValueKind.Date ? PivotMaterializedDateCaption(key, plan.SourceDateSystem) : key.Text;
                            label = PivotMaterializedText(dateRowLayout && axis.Layout.HasValues && rowValuesPosition > 0
                                ? text + " " + (measures[entry.Measure].Name?.Value ?? "") : text + " Total");
                        }
                        put(level, label, key.Kind == PivotFieldValueKind.Date && !subtotalLabel);
                    }
                }
            }
            var rowRealDepths = new int[rows.Layout.Fields.Length];
            for (int level = 0, depth = 0; level < rowRealDepths.Length; level++) {
                if (rows.Layout.Fields[level] >= 0) depth++;
                rowRealDepths[level] = depth;
            }
            PivotHierarchyEntry? previousDataEntry = null;
            int[]? previousDataKeys = null, previousRowKeys = null;
            for (int row = 0; row < rows.Entries.Count; row++) {
                token.ThrowIfCancellationRequested();
                int position = dataRow + row;
                var entry = rows.Entries[row];
                int[] rowKeys = MaterializedHierarchyKeys(entry.Node);
                bool firstMeasureRow = previousDataEntry == null || previousDataEntry.Value.Measure != entry.Measure
                    || rowValuesPosition > 0 && !MaterializedKeyPrefixEquals(previousDataKeys!, rowKeys, rowValuesPosition);
                Labels(rows, entry, rowKeys, dateHierarchy, firstMeasureRow, (level, value, date) => {
                    int realDepth = rowRealDepths[level];
                    if (dateHierarchy && rows.Layout.Fields[level] >= 0
                        && (realDepth < rows.Layout.RealFields.Length || rowValuesPosition > level)
                        && rows.Entries[row].Type == ItemValues.Data && row > 0
                        && rows.Entries[row - 1].Type == ItemValues.Data
                        && MaterializedKeyPrefixEquals(previousRowKeys!, rowKeys, realDepth)) return;
                    values[position, level] = value;
                    if (date) plan.DateCells.Add((position, level));
                });
                previousRowKeys = rowKeys;
                if (entry.Type == ItemValues.Data) {
                    previousDataEntry = entry;
                    previousDataKeys = rowKeys;
                }
            }
            var columnRealDepths = new int[columns.Layout.Fields.Length];
            for (int level = 0, depth = 0; level < columnRealDepths.Length; level++) {
                if (columns.Layout.Fields[level] >= 0) depth++;
                columnRealDepths[level] = depth;
            }
            int[]? previousColumnKeys = null;
            for (int column = 0; column < columns.Entries.Count; column++) {
                int position = dataColumn + column;
                int offset = columns.Layout.RealFields.Length > 0 ? 1 : 0;
                var entry = columns.Entries[column];
                int[] columnKeys = MaterializedHierarchyKeys(entry.Node);
                Labels(columns, entry, columnKeys, false, true, (level, value, date) => {
                    int realDepth = columnRealDepths[level];
                    if (dateColumnHierarchy && columns.Layout.Fields[level] >= 0
                        && realDepth < columns.Layout.RealFields.Length
                        && entry.Type == ItemValues.Data && column > 0
                        && columns.Entries[column - 1].Type == ItemValues.Data
                        && MaterializedKeyPrefixEquals(previousColumnKeys!, columnKeys, realDepth)) return;
                    values[level + offset, position] = value;
                    if (date) plan.DateCells.Add((level + offset, position));
                });
                previousColumnKeys = columnKeys;
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
            Location location = definition.Location ?? new Location();
            location.Reference = $"{A1.CellReference(plan.Top, plan.Left)}:{A1.CellReference(plan.Bottom, plan.Right)}";
            location.FirstHeaderRow = columns.Layout.RealFields.Length > 0 || !columns.Layout.HasValues ? 1U : 0U;
            location.FirstDataRow = (uint)dataRow;
            location.FirstDataColumn = (uint)dataColumn;
            definition.Location = location;
            definition.Compact = false;
            definition.CompactData = false;
            definition.OutlineData = false;
            definition.DataOnRows = rows.Layout.HasValues;
            NormalizeMaterializedPivotFields(definition, maps, rows, visibility);
            NormalizeMaterializedPivotFields(definition, maps, columns, visibility);
            NormalizeMaterializedPageFields(definition, maps, visibility);
            definition.RowItems = new RowItems { Count = (uint)rows.Entries.Count };
            definition.ColumnItems = new ColumnItems { Count = (uint)columns.Entries.Count };
            foreach (var entry in rows.Entries) definition.RowItems.AppendChild(CreateMaterializedHierarchyItem(rows, entry));
            foreach (var entry in columns.Entries) definition.ColumnItems.AppendChild(CreateMaterializedHierarchyItem(columns, entry));
        }

        private static void NormalizeMaterializedPivotFields(PivotTableDefinition definition, IReadOnlyList<PivotFieldValues> maps,
            PivotHierarchyAxis axis, PivotMaterializationVisibility visibility) {
            var fields = definition.PivotFields!.Elements<PivotField>().ToArray();
            for (int depth = 0; depth < axis.Layout.RealFields.Length; depth++) {
                int field = axis.Layout.RealFields[depth];
                bool subtotal = axis.Subtotals[depth];
                fields[field].Items = CreateMaterializedFilteredItems(maps[field],
                    visibility.Hidden.TryGetValue(field, out var hidden) ? hidden : null, false, subtotal);
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

        private static bool MaterializedKeyPrefixEquals(int[] left, int[] right, int count) {
            if (left.Length < count || right.Length < count) return false;
            for (int index = 0; index < count; index++) if (left[index] != right[index]) return false;
            return true;
        }

        private static ExcelCellData PivotMaterializedText(string text) => new(ExcelCellDataKind.Text, text);
        private static string PivotMaterializedDateCaption(PivotFieldValue key, ExcelDateSystem dateSystem) {
            double serial = ExcelPivotCacheDateCodec.ToSerial(key.Date!.Value, dateSystem);
            return FormatWorksheetDateText(serial, ExcelDateSystemConverter.FromSerial(serial, dateSystem), "yyyy-MM-dd HH:mm:ss", dateSystem);
        }
        private static ExcelCellData PivotMaterializedKey(PivotFieldValue key, ExcelDateSystem dateSystem) => key.Kind switch {
            PivotFieldValueKind.Blank => PivotMaterializedText("(blank)"),
            PivotFieldValueKind.Boolean => new(ExcelCellDataKind.Boolean, key.Boolean),
            PivotFieldValueKind.Number => new(ExcelCellDataKind.Number, key.Number),
            PivotFieldValueKind.Date => new(ExcelCellDataKind.Number, ExcelPivotCacheDateCodec.ToSerial(key.Date!.Value, dateSystem)),
            PivotFieldValueKind.Error => new(ExcelCellDataKind.Error, key.Text, cachedText: key.Text),
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
