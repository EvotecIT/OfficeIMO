using System.Diagnostics.CodeAnalysis;
using OfficeIMO.Data;

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
            ExcelTableStyle style = ExcelTableStyle.TableStyleMedium9,
            bool autoFilter = true,
            bool freezeHeaderRow = true,
            System.Action<TableVisualOptions>? visuals = null) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            string? maxColumnsGuidance = opts.MaxColumns >= A1.MaxColumns
                ? $"Select fewer columns or split the data across multiple worksheets; Excel's {A1.MaxColumns}-column worksheet limit cannot be overridden."
                : null;
            opts.MaxColumns = System.Math.Min(opts.MaxColumns, A1.MaxColumns);
            var flattener = new ObjectFlattener();

            ObjectTableProjection projection = flattener.FlattenRows(
                items,
                opts,
                "TableFrom",
                headerRowCount: 1,
                enforceEmptyProjectionLimits: false,
                columnLimitGuidance: maxColumnsGuidance);
            IReadOnlyList<System.Collections.Generic.Dictionary<string, object?>> rows = projection.Rows;
            if (rows.Count == 0) {
                Sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            IReadOnlyList<string> paths = opts.Columns == null
                ? projection.Columns.OrderBy(path => path, System.StringComparer.Ordinal).ToList()
                : projection.Columns;

            // If we still have no columns (e.g., row type exposes fields but no public properties),
            // degrade gracefully rather than producing an invalid table definition.
            if (paths.Count == 0) {
                Sheet.Cell(_row, 1, "(no tabular columns for row type)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(System.Math.Max(1, (rows.Count + 1) * System.Math.Max(1, paths.Count)));
            var headersT = BuildTransformedHeaders(paths, opts);
            EnsureUniqueTableHeaders(headersT);
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

            int lastRow = _row - 1;
            string start = $"A{headerRow}";
            string end = ColumnLetter(paths.Count) + lastRow.ToString();
            string range = start + ":" + end;

            var tableName = title ?? "Table";
            Sheet.AddTable(range, hasHeader: true, name: tableName, style: style, includeAutoFilter: autoFilter);
            return CompleteTable(range, paths, headersT, headerRow, lastRow, style, freezeHeaderRow, visuals);
        }

    }
}
