using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Simple multi-column layout for placing lightweight blocks side by side.
    /// </summary>
    public sealed partial class SheetComposer
    {
        /// <summary>
        /// Column-scoped composer used by <see cref="Columns"/> to
        /// write simple sections into a single column while keeping track of vertical position.
        /// </summary>
        public sealed class ColumnComposer
        {
            private readonly ExcelSheet _sheet;
            private readonly SheetTheme _theme;
            private readonly int _baseCol;
            private readonly int _startRow;
            private int _row;
            private int _maxColUsed;
            private int? _maxTableColumns;
            private OverflowMode _overflowMode = OverflowMode.Throw;

            internal ColumnComposer(ExcelSheet sheet, SheetTheme theme, int startRow, int baseCol)
            { _sheet = sheet; _theme = theme; _startRow = startRow; _row = startRow; _baseCol = baseCol; _maxColUsed = baseCol; }

            /// <summary>Total number of rows consumed by this column since it was created.</summary>
            public int RowsUsed => _row - _startRow;

            /// <summary>Total number of sheet columns used by this column content (>=1).</summary>
            public int ColumnsUsed => Math.Max(1, _maxColUsed - _baseCol + 1);

            internal void SetGridConstraints(int columnWidth, OverflowMode overflow)
            {
                _maxTableColumns = Math.Max(1, columnWidth);
                _overflowMode = overflow;
            }

            /// <summary>Adds vertical space.</summary>
            /// <param name="rows">Number of rows to skip (minimum 1).</param>
            public ColumnComposer Spacer(int rows = 1) { _row += Math.Max(1, rows); return this; }

            /// <summary>Writes a section header cell using the current theme's section style.</summary>
            public ColumnComposer Section(string text)
            {
                _sheet.Cell(_row, _baseCol, text);
                _sheet.CellBold(_row, _baseCol, true);
                _sheet.CellBackground(_row, _baseCol, _theme.SectionHeaderFillHex);
                if (_baseCol > 0) _maxColUsed = Math.Max(_maxColUsed, _baseCol);
                _row++;
                return this;
            }

            /// <summary>Writes a single paragraph cell and advances the row.</summary>
            public ColumnComposer Paragraph(string text)
            {
                if (!string.IsNullOrEmpty(text)) { _sheet.Cell(_row, _baseCol, text); _row++; }
                if (_baseCol > 0) _maxColUsed = Math.Max(_maxColUsed, _baseCol);
                return this;
            }

            /// <summary>Writes a bullet for each item (• text) in this column.</summary>
            public ColumnComposer BulletedList(IEnumerable<string> items)
            {
                if (items == null) return this;
                foreach (var item in items) { _sheet.Cell(_row, _baseCol, $"• {item}"); _row++; }
                if (_baseCol > 0) _maxColUsed = Math.Max(_maxColUsed, _baseCol);
                return this;
            }

            /// <summary>Writes a two-column key/value row with styled key cell.</summary>
            public ColumnComposer KeyValue(string key, object? value)
            {
                _sheet.Cell(_row, _baseCol, key);
                _sheet.CellBold(_row, _baseCol, true);
                _sheet.CellBackground(_row, _baseCol, _theme.KeyFillHex);
                _sheet.Cell(_row, _baseCol + 1, value ?? string.Empty);
                _maxColUsed = Math.Max(_maxColUsed, _baseCol + 1);
                _row++;
                return this;
            }

            /// <summary>Writes multiple key/value rows in order.</summary>
            public ColumnComposer KeyValues(IEnumerable<(string Key, object? Value)> pairs)
            {
                if (pairs == null) return this;
                foreach (var (k, v) in pairs) KeyValue(k, v);
                return this;
            }

            /// <summary>
            /// Renders a table inside this column starting at the current row and returns the A1 range.
            /// This variant avoids freezing panes to prevent conflicts when multiple tables are placed per sheet.
            /// </summary>
            public string TableFrom<T>(IEnumerable<T> items, string? title = null,
                System.Action<ObjectFlattenerOptions>? configure = null,
                TableStyle style = TableStyle.TableStyleMedium9,
                bool autoFilter = true,
                System.Action<TableVisualOptions>? visuals = null)
            {
                if (!string.IsNullOrWhiteSpace(title)) Section(title!);

                var list = items?.ToList() ?? new List<T>();
                if (list.Count == 0)
                {
                    _sheet.Cell(_row, _baseCol, "(no data)");
                    _row++;
                    return $"{SheetComposer.ColumnLetter(_baseCol)}{_row-1}:{SheetComposer.ColumnLetter(_baseCol)}{_row-1}";
                }

                var opts = new ObjectFlattenerOptions();
                configure?.Invoke(opts);
                var flattener = new ObjectFlattener();

                var rows = new List<Dictionary<string, object?>>();
                foreach (var item in list) rows.Add(flattener.Flatten(item, opts));

                var paths = opts.Columns?.ToList() ?? rows.SelectMany(r => r.Keys)
                    .Where(k => !string.IsNullOrWhiteSpace(k))
                    .Distinct(System.StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, System.StringComparer.Ordinal)
                    .ToList();
                if (paths.Count == 0)
                {
                    _sheet.Cell(_row, _baseCol, "(no tabular columns)");
                    _row++;
                    return $"{SheetComposer.ColumnLetter(_baseCol)}{_row-1}:{SheetComposer.ColumnLetter(_baseCol)}{_row-1}";
                }

                int headerRow = _row;
                // Enforce fixed-grid overflow behavior if configured
                List<string> effPaths = paths;
                bool summarize = false;
                if (_maxTableColumns.HasValue && paths.Count > _maxTableColumns.Value)
                {
                    if (_overflowMode == OverflowMode.Throw)
                        throw new InvalidOperationException($"Table has {paths.Count} columns but only {_maxTableColumns.Value} fit in the fixed grid. Increase columnWidth or use ColumnsAdaptive(...).");
                    if (_overflowMode == OverflowMode.Shrink)
                    {
                        effPaths = paths.Take(_maxTableColumns.Value).ToList();
                        try { _sheet.EffectiveExecution.ReportInfo($"[Columns Shrink] Sheet='{_sheet.Name}', baseCol={_baseCol}, kept={effPaths.Count}, dropped={paths.Count - effPaths.Count}"); } catch { }
                    }
                    else if (_overflowMode == OverflowMode.Summarize)
                    {
                        int keep = Math.Max(1, _maxTableColumns.Value - 1);
                        if (keep <= 0)
                        {
                            effPaths = new List<string> { "__More__" };
                        }
                        else
                        {
                            effPaths = paths.Take(keep).ToList();
                            effPaths.Add("__More__");
                        }
                        summarize = true;
                        try { _sheet.EffectiveExecution.ReportInfo($"[Columns Summarize] Sheet='{_sheet.Name}', baseCol={_baseCol}, kept={(effPaths.Contains("__More__") ? effPaths.Count - 1 : effPaths.Count)}, summarized={paths.Count - (effPaths.Contains("__More__") ? effPaths.Count - 1 : effPaths.Count)}"); } catch { }
                    }
                }

                var headersT = effPaths.Select(p => p == "__More__" ? "More" : SheetComposer.TransformHeader(p, opts)).ToList();
                var usedHeaders = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < headersT.Count; i++)
                {
                    string baseName = string.IsNullOrWhiteSpace(headersT[i]) ? $"Column{i+1}" : headersT[i];
                    string candidate = baseName; int suffix = 2;
                    while (!usedHeaders.Add(candidate)) candidate = $"{baseName} ({suffix++})";
                    headersT[i] = candidate;
                }
                for (int i = 0; i < headersT.Count; i++)
                {
                    _sheet.Cell(headerRow, _baseCol + i, headersT[i]);
                    _sheet.CellBold(headerRow, _baseCol + i, true);
                    _sheet.CellBackground(headerRow, _baseCol + i, _theme.KeyFillHex);
                }
                _row++;
                foreach (var dict in rows)
                {
                    if (summarize)
                    {
                        string more = string.Empty;
                        if (_maxTableColumns.HasValue && paths.Count > _maxTableColumns.Value)
                        {
                            int keep = Math.Max(1, _maxTableColumns.Value - 1);
                            var omitted = paths.Skip(keep);
                            var parts = new System.Collections.Generic.List<string>();
                            foreach (var p in omitted)
                            {
                                dict.TryGetValue(p, out var v);
                                var label = SheetComposer.TransformHeader(p, opts);
                                parts.Add(string.Concat(label, "=", v?.ToString() ?? string.Empty));
                            }
                            more = string.Join("; ", parts);
                        }
                        for (int i = 0; i < effPaths.Count; i++)
                        {
                            object? val;
                            if (effPaths[i] == "__More__") val = more; else { dict.TryGetValue(effPaths[i], out val); }
                            _sheet.Cell(_row, _baseCol + i, val ?? string.Empty);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < effPaths.Count; i++)
                        {
                            dict.TryGetValue(effPaths[i], out var val);
                            _sheet.Cell(_row, _baseCol + i, val ?? string.Empty);
                        }
                    }
                    _row++;
                }

                int lastRow = _row - 1;
                string start = SheetComposer.ColumnLetter(_baseCol) + headerRow.ToString();
                string end = SheetComposer.ColumnLetter(_baseCol + headersT.Count - 1) + lastRow.ToString();
                string range = start + ":" + end;

                string tableName = title ?? "Table";
                _sheet.AddTable(range, hasHeader: true, name: tableName, style: style, includeAutoFilter: autoFilter);

                var viz = new TableVisualOptions(); visuals?.Invoke(viz);
                try
                {
                    for (int i = 0; i < headersT.Count; i++)
                    {
                        string hdr = headersT[i];
                        string colRange = $"{SheetComposer.ColumnLetter(_baseCol + i)}{headerRow + 1}:{SheetComposer.ColumnLetter(_baseCol + i)}{_row - 1}";
                        if (viz.NumericColumnFormats.TryGetValue(hdr, out var fmt))
                            _sheet.ColumnStyleByHeader(hdr).NumberFormat(fmt);
                        else if (viz.NumericColumnDecimals.TryGetValue(hdr, out var dec))
                            _sheet.ColumnStyleByHeader(hdr).Number(dec);
                        if (viz.DataBars.TryGetValue(hdr, out var color))
                            _sheet.AddConditionalDataBar(colRange, color);
                        if (viz.IconSets.TryGetValue(hdr, out var icon))
                            _sheet.AddConditionalIconSet(colRange, icon.IconSet, icon.ShowValue, icon.ReverseOrder, icon.PercentThresholds, icon.NumberThresholds);
                        else if (viz.IconSetColumns.Contains(hdr))
                            _sheet.AddConditionalIconSet(colRange);
                    }
                }
                catch { }

                // Track used width for adaptive column placement
                int usedCols = Math.Max(1, headersT.Count);
                _maxColUsed = Math.Max(_maxColUsed, _baseCol + usedCols - 1);
                return range;
            }

            /// <summary>
            /// Splits the current column region into horizontal sub-columns, rendering each via a nested ColumnComposer.
            /// Advances the parent column by the tallest child block and applies a spacer afterwards.
            /// </summary>
            /// <param name="count">Number of sub-columns to create.</param>
            /// <param name="configure">Callback receiving the nested ColumnComposer instances.</param>
            /// <param name="columnWidth">Width per sub-column in grid columns.</param>
            /// <param name="gutter">Spacing between sub-columns in grid columns.</param>
            public ColumnComposer Columns(int count, Action<ColumnComposer[]> configure, int columnWidth = 3, int gutter = 1, OverflowMode overflow = OverflowMode.Throw)
            {
                if (count <= 1) return this;
                int startRow = _row;
                var cols = new ColumnComposer[count];
                int baseCol = _baseCol;
                for (int i = 0; i < count; i++)
                {
                    cols[i] = new ColumnComposer(_sheet, _theme, startRow, baseCol);
                    cols[i].SetGridConstraints(columnWidth, overflow);
                    baseCol += columnWidth + gutter;
                }
                configure?.Invoke(cols);
                int maxRows = 0;
                foreach (var c in cols)
                {
                    if (c.RowsUsed > maxRows) maxRows = c.RowsUsed;
                }
                _row = startRow + maxRows;
                return Spacer();
            }
        }

        /// <summary>
        /// Places N columns side-by-side starting at the current row. Each action receives a ColumnComposer
        /// scoped to its own column. The main composer advances to the maximum height used by the columns.
        /// </summary>
        /// <param name="count">Number of columns (2–4 recommended).</param>
        /// <param name="configure">Callback that receives an array of ColumnComposer objects.</param>
        /// <param name="columnWidth">Width per column in grid columns (for relative positioning only).</param>
        /// <param name="gutter">Spacing between columns in grid columns.</param>
        public SheetComposer Columns(int count, Action<ColumnComposer[]> configure, int columnWidth = 3, int gutter = 1, OverflowMode overflow = OverflowMode.Throw)
        {
            if (count <= 1) return this;
            int startRow = _row;
            var cols = new ColumnComposer[count];
            int baseCol = 1;
            for (int i = 0; i < count; i++)
            {
                cols[i] = new ColumnComposer(Sheet, _theme, startRow, baseCol);
                cols[i].SetGridConstraints(columnWidth, overflow);
                baseCol += columnWidth + gutter;
            }
            configure?.Invoke(cols);
            int maxRows = 0; foreach (var c in cols) if (c.RowsUsed > maxRows) maxRows = c.RowsUsed;
            _row = startRow + maxRows;
            return Spacer();
        }

        /// <summary>
        /// Places columns left-to-right starting at the current row using adaptive widths derived from each column's content.
        /// Each column is rendered at the computed base column and the next column starts at (previous end + 1 + gutter).
        /// </summary>
        public SheetComposer ColumnsAdaptive(IReadOnlyList<Action<ColumnComposer>> builders, int gutter = 1)
        {
            if (builders == null || builders.Count == 0) return this;
            int startRow = _row;
            int baseCol = 1;
            int maxRows = 0;
            foreach (var b in builders)
            {
                if (b == null) continue;
                var col = new ColumnComposer(Sheet, _theme, startRow, baseCol);
                b(col);
                maxRows = Math.Max(maxRows, col.RowsUsed);
                baseCol += col.ColumnsUsed + Math.Max(0, gutter);
            }
            _row = startRow + maxRows;
            return Spacer();
        }

        /// <summary>
        /// Renders multiple rows of adaptive columns. Each inner list is a left-to-right band; rows stack vertically.
        /// </summary>
        public SheetComposer ColumnsAdaptiveRows(IReadOnlyList<IReadOnlyList<Action<ColumnComposer>>> rows, int gutter = 1)
        {
            if (rows == null || rows.Count == 0) return this;
            foreach (var row in rows)
            {
                ColumnsAdaptive(row, gutter);
            }
            return this;
        }

        /// <summary>
        /// Renders a sequence of column blocks left-to-right using a fixed number of columns per band.
        /// Each band uses <see cref="Columns(int, Action{ColumnComposer[]}, int, int)"/> under the hood and advances
        /// the composer by the tallest block in that band.
        /// </summary>
        public SheetComposer FlowColumns(IReadOnlyList<IReadOnlyList<Action<ColumnComposer>>> columnGroups, int columnWidth = 12, int gutter = 2)
        {
            if (columnGroups == null || columnGroups.Count == 0)
            {
                return this;
            }

            int columns = columnGroups.Count;
            int index = 0;
            while (true)
            {
                bool any = false;
                for (int i = 0; i < columns; i++)
                {
                    var group = columnGroups[i];
                    if (group != null && index < group.Count)
                    {
                        any = true;
                        break;
                    }
                }

                if (!any)
                {
                    break;
                }

                Columns(columns, cols =>
                {
                    for (int i = 0; i < columns; i++)
                    {
                        var group = columnGroups[i];
                        if (group != null && index < group.Count)
                        {
                            group[index](cols[i]);
                        }
                    }
                }, columnWidth, gutter);

                index++;
            }

            return this;
        }
    }
}
