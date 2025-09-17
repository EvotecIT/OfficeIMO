using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Table rendering for SheetComposer.
    /// </summary>
    public sealed partial class SheetComposer
    {
        /// <summary>
        /// Flattens a sequence of objects into a table and renders it with a header row.
        /// Returns the A1 range used for the table.
        /// </summary>
        public string TableFrom<T>(IEnumerable<T> items, string? title = null,
            System.Action<ObjectFlattenerOptions>? configure = null,
            TableStyle style = TableStyle.TableStyleMedium9,
            bool autoFilter = true,
            bool freezeHeaderRow = true,
            System.Action<TableVisualOptions>? visuals = null)
        {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var data = items?.ToList() ?? new List<T>();
            if (data.Count == 0)
            {
                Sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row-1}:A{_row-1}";
            }

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            var flattener = new ObjectFlattener();

            var rows = new List<System.Collections.Generic.Dictionary<string, object?>>();
            foreach (var item in data)
                rows.Add(flattener.Flatten(item, opts));

            var paths = opts.Columns?.ToList();
            if (paths == null)
            {
                paths = rows.SelectMany(r => r.Keys)
                            .Where(k => !string.IsNullOrWhiteSpace(k))
                            .Distinct(System.StringComparer.OrdinalIgnoreCase)
                            .OrderBy(s => s, System.StringComparer.Ordinal)
                            .ToList();
                // Apply selection filters (Ignore/Exclude/Include) then ordering (Pinned/Priority)
                paths = OfficeIMO.Excel.ObjectFlattener.ApplySelection(paths, opts);
                paths = OfficeIMO.Excel.ObjectFlattener.ApplyOrdering(paths, opts);
            }

            // If we still have no columns (e.g., row type exposes fields but no public properties),
            // degrade gracefully rather than producing an invalid table definition.
            if (paths.Count == 0)
            {
                Sheet.Cell(_row, 1, "(no tabular columns for row type)");
                _row++;
                return $"A{_row-1}:A{_row-1}";
            }

            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(System.Math.Max(1, (rows.Count + 1) * System.Math.Max(1, paths.Count)));
            var headersT = paths.Select(p => TransformHeader(p, opts)).ToList();
            // De-duplicate header captions to avoid Excel silently renaming duplicates
            var usedHeaders = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headersT.Count; i++)
            {
                string baseName = string.IsNullOrWhiteSpace(headersT[i]) ? $"Column{i+1}" : headersT[i];
                string candidate = baseName;
                int suffix = 2;
                while (!usedHeaders.Add(candidate))
                {
                    candidate = $"{baseName} ({suffix++})";
                }
                headersT[i] = candidate;
            }
            var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headersT.Count; i++)
            {
                var h = headersT[i];
                if (!seen.Add(h))
                {
                    int n = 2; string candidate;
                    do { candidate = h + "_" + n++; } while (!seen.Add(candidate));
                    headersT[i] = candidate;
                }
                cells.Add((_row, i + 1, headersT[i]));
            }
            _row++;

            foreach (var dict in rows)
            {
                for (int i = 0; i < paths.Count; i++)
                {
                    dict.TryGetValue(paths[i], out var val);
                    cells.Add((_row, i + 1, val ?? string.Empty));
                }
                _row++;
            }
            Sheet.CellValues(cells);
            for (int i = 0; i < paths.Count; i++)
            {
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
            if (viz.FreezeHeaderRow) { try { Sheet.Freeze(topRows: headerRow, leftCols: 0); } catch { } }

            try
            {
                var headers = headersT; int startCol = 1;
                for (int i = 0; i < headers.Count; i++)
                {
                    string hdr = headers[i];
                    string colRange = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";

                    if (viz.NumericColumnFormats.TryGetValue(hdr, out var fmt))
                    {
                        if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                            Sheet.ColumnStyleByHeader(hdr).NumberFormat(fmt);
                    }
                    else if (viz.NumericColumnDecimals.TryGetValue(hdr, out var dec))
                    {
                        if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                            Sheet.ColumnStyleByHeader(hdr).Number(dec);
                    }

                    if (viz.DataBars.TryGetValue(hdr, out var color))
                        try { Sheet.AddConditionalDataBar(colRange, color); } catch { }

                    if (viz.IconSets.TryGetValue(hdr, out var iconOpts))
                        try { Sheet.AddConditionalIconSet(colRange, iconOpts.IconSet, iconOpts.ShowValue, iconOpts.ReverseOrder, iconOpts.PercentThresholds, iconOpts.NumberThresholds); } catch { }
                    else if (viz.IconSetColumns.Contains(hdr))
                        try { Sheet.AddConditionalIconSet(colRange); } catch { }

                    if (viz.TextBackgrounds.TryGetValue(hdr, out var map))
                    {
                        if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                        {
                            Sheet.ColumnStyleByHeader(hdr).BackgroundByTextMap(map);
                        }
                        else
                        {
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                                if (Sheet.TryGetCellText(r, startCol + i, out var t) && t != null && map.TryGetValue(t, out var colorHex))
                                    Sheet.CellBackground(r, startCol + i, colorHex);
                        }
                    }
                    if (viz.BoldByText.TryGetValue(hdr, out var boldSet))
                    {
                        if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                        {
                            Sheet.ColumnStyleByHeader(hdr).BoldByTextSet(boldSet);
                        }
                        else
                        {
                            var setCI = new System.Collections.Generic.HashSet<string>(boldSet, System.StringComparer.OrdinalIgnoreCase);
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                                if (Sheet.TryGetCellText(r, startCol + i, out var t) && !string.IsNullOrEmpty(t) && setCI.Contains(t))
                                    Sheet.CellBold(r, startCol + i, true);
                        }
                    }
                }

                if (viz.AutoFormatDynamicCollections)
                {
                    for (int i = 0; i < paths.Count; i++)
                    {
                        if (paths[i].Contains('.'))
                        {
                            var hdr = headers[i];
                            if (Sheet.TryGetColumnIndexByHeader(hdr, out _))
                                Sheet.ColumnStyleByHeader(hdr).Number(viz.AutoFormatDecimals);
                            string colRangeAuto = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";
                            try { Sheet.AddConditionalDataBar(colRangeAuto, viz.AutoFormatDataBarColor); } catch { }
                        }
                    }
                }
            }
            catch { }
            Spacer();
            return range;
        }
    }
}
