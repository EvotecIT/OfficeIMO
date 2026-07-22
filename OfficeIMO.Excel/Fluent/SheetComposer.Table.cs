using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Table rendering for SheetComposer.
    /// </summary>
    public sealed partial class SheetComposer {
        /// <summary>
        /// Flattens a sequence of objects into a table and renders it with a header row.
        /// Returns the A1 range used for the table.
        /// </summary>
        public string TableFrom<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IEnumerable<T> items, string? title = null,
            System.Action<ObjectFlattenerOptions>? configure = null,
            TableStyle style = TableStyle.TableStyleMedium9,
            bool autoFilter = true,
            bool freezeHeaderRow = true,
            System.Action<TableVisualOptions>? visuals = null) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            var flattener = new ObjectFlattener();

            var rows = FlattenTableRows(items, flattener, opts);
            if (rows.Count == 0) {
                Sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            var paths = opts.Columns?.ToList();
            if (paths == null) {
                paths = rows.SelectMany(r => r.Keys)
                            .Where(k => !string.IsNullOrWhiteSpace(k))
                            .Distinct(System.StringComparer.OrdinalIgnoreCase)
                            .OrderBy(s => s, System.StringComparer.Ordinal)
                            .ToList();
                // Apply selection filters (Ignore/Exclude/Include) then ordering (Pinned/Priority)
                paths = flattener.ResolvePaths(paths, opts);
            }

            // If we still have no columns (e.g., row type exposes fields but no public properties),
            // degrade gracefully rather than producing an invalid table definition.
            if (paths.Count == 0) {
                Sheet.Cell(_row, 1, "(no tabular columns for row type)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(System.Math.Max(1, (rows.Count + 1) * System.Math.Max(1, paths.Count)));
            var headersT = paths.Select(p => TransformHeader(p, opts)).ToList();
            // De-duplicate header captions to avoid Excel silently renaming duplicates
            var usedHeaders = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headersT.Count; i++) {
                string baseName = string.IsNullOrWhiteSpace(headersT[i]) ? $"Column{i + 1}" : headersT[i];
                string candidate = baseName;
                int suffix = 2;
                while (!usedHeaders.Add(candidate)) {
                    candidate = $"{baseName} ({suffix++})";
                }
                headersT[i] = candidate;
            }
            for (int i = 0; i < headersT.Count; i++) {
                cells.Add((_row, i + 1, headersT[i]));
            }
            _row++;

            foreach (var dict in rows) {
                for (int i = 0; i < paths.Count; i++) {
                    dict.TryGetValue(paths[i], out var val);
                    cells.Add((_row, i + 1, val ?? string.Empty));
                }
                _row++;
            }
            Sheet.CellValues(cells);
            for (int i = 0; i < paths.Count; i++) {
                Sheet.CellBold(headerRow, i + 1, true);
                Sheet.CellBackground(headerRow, i + 1, _theme.KeyFillHex);
            }

            int lastRow = _row - 1;
            string start = $"A{headerRow}";
            string end = ColumnLetter(paths.Count) + lastRow.ToString();
            string range = start + ":" + end;

            var tableName = title ?? "Table";
            Sheet.AddTable(range, hasHeader: true, name: tableName, style: style, includeAutoFilter: autoFilter);

            var viz = new TableVisualOptions();
            viz.FreezeHeaderRow = freezeHeaderRow; visuals?.Invoke(viz);
            Sheet.SetTableStyle(range, style, viz.ShowFirstColumn, viz.ShowLastColumn, viz.ShowRowStripes, viz.ShowColumnStripes);
            if (viz.FreezeHeaderRow) Sheet.Freeze(topRows: headerRow, leftCols: 0);

            var headers = headersT; int startCol = 1;
            for (int i = 0; i < headers.Count; i++) {
                string hdr = headers[i];
                string colRange = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";

                if (viz.NumericColumnFormats.TryGetValue(hdr, out var fmt)) {
                    if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                        Sheet.ColumnStyleByHeader(hdr).NumberFormat(fmt);
                } else if (viz.NumericColumnDecimals.TryGetValue(hdr, out var dec)) {
                    if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                        Sheet.ColumnStyleByHeader(hdr).Number(dec);
                }

                if (viz.DataBars.TryGetValue(hdr, out var color))
                    Sheet.AddConditionalDataBar(colRange, color);

                if (viz.IconSets.TryGetValue(hdr, out var iconOpts))
                    Sheet.AddConditionalIconSet(colRange, iconOpts.IconSet, iconOpts.ShowValue, iconOpts.ReverseOrder, iconOpts.PercentThresholds, iconOpts.NumberThresholds);
                else if (viz.IconSetColumns.Contains(hdr))
                    Sheet.AddConditionalIconSet(colRange);

                if (viz.TextBackgrounds.TryGetValue(hdr, out var map)) {
                    if (Sheet.TryGetColumnIndexByHeader(hdr, out _)) {
                        Sheet.ColumnStyleByHeader(hdr).BackgroundByTextMap(map);
                    } else {
                        for (int r = headerRow + 1; r <= _row - 1; r++)
                            if (Sheet.TryGetCellText(r, startCol + i, out var t) && t != null && map.TryGetValue(t, out var colorHex))
                                Sheet.CellBackground(r, startCol + i, colorHex);
                    }
                }
                if (viz.BoldByText.TryGetValue(hdr, out var boldSet)) {
                    if (Sheet.TryGetColumnIndexByHeader(hdr, out _)) {
                        Sheet.ColumnStyleByHeader(hdr).BoldByTextSet(boldSet);
                    } else {
                        var setCI = new System.Collections.Generic.HashSet<string>(boldSet, System.StringComparer.OrdinalIgnoreCase);
                        for (int r = headerRow + 1; r <= _row - 1; r++)
                            if (Sheet.TryGetCellText(r, startCol + i, out var t) && !string.IsNullOrEmpty(t) && setCI.Contains(t))
                                Sheet.CellBold(r, startCol + i, true);
                    }
                }
            }

            if (viz.AutoFormatDynamicCollections) {
                for (int i = 0; i < paths.Count; i++) {
                    if (paths[i].Contains('.')) {
                        var hdr = headers[i];
                        if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                            Sheet.ColumnStyleByHeader(hdr).Number(viz.AutoFormatDecimals);
                        string colRangeAuto = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";
                        Sheet.AddConditionalDataBar(colRangeAuto, viz.AutoFormatDataBarColor);
                    }
                }
            }
            Spacer();
            return range;
        }

        private static List<System.Collections.Generic.Dictionary<string, object?>> FlattenTableRows<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
            IEnumerable<T>? items,
            ObjectFlattener flattener,
            ObjectFlattenerOptions options) {
            if (items == null) {
                return new List<System.Collections.Generic.Dictionary<string, object?>>();
            }

            if (items is IReadOnlyList<T> indexedItems) {
                var rows = new List<System.Collections.Generic.Dictionary<string, object?>>(indexedItems.Count);
                for (int i = 0; i < indexedItems.Count; i++) {
                    rows.Add(flattener.Flatten(indexedItems[i], options));
                }

                return rows;
            }

            int capacity = items is IReadOnlyCollection<T> readOnlyCollection
                ? readOnlyCollection.Count
                : items is ICollection<T> collection ? collection.Count : 0;
            var materializedRows = capacity > 0
                ? new List<System.Collections.Generic.Dictionary<string, object?>>(capacity)
                : new List<System.Collections.Generic.Dictionary<string, object?>>();
            foreach (var item in items) {
                materializedRows.Add(flattener.Flatten(item, options));
            }

            return materializedRows;
        }
    }
}
