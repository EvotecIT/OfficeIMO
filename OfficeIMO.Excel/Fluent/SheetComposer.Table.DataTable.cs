using System.Data;
using System.IO;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.Fluent {
    public sealed partial class SheetComposer {
        /// <summary>
        /// Renders a fixed-schema <see cref="DataTable"/> as an editable Excel table without generic object flattening.
        /// The DataTable schema order is preserved unless the projection configuration selects another order.
        /// </summary>
        /// <param name="table">Source table whose rows and schema are already materialized.</param>
        /// <param name="title">Optional section title and table name.</param>
        /// <param name="configure">
        /// Optional column selection, exclusion, ordering, and header configuration. Generic object-expansion and
        /// cell-materialization limits are not used by this fixed-schema overload.
        /// </param>
        /// <param name="style">Built-in Excel table style.</param>
        /// <param name="autoFilter">Whether table filter buttons are included.</param>
        /// <param name="freezeHeaderRow">Whether rows through the table header are frozen.</param>
        /// <param name="visuals">Optional table visual configuration.</param>
        /// <returns>The A1 range occupied by the table.</returns>
        public string TableFrom(
            DataTable table,
            string? title = null,
            Action<ObjectFlattenerOptions>? configure = null,
            ExcelTableStyle style = ExcelTableStyle.TableStyleMedium9,
            bool autoFilter = true,
            bool freezeHeaderRow = true,
            Action<TableVisualOptions>? visuals = null) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            if (table.Rows.Count == 0) {
                Sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            int configuredMaxColumns = options.MaxColumns;
            if (configuredMaxColumns <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxColumns), "MaxColumns must be greater than zero.");
            }
            options.MaxColumns = Math.Min(options.MaxColumns, A1.MaxColumns);

            var sourceColumnNames = table.Columns
                .Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .ToArray();
            var flattener = new ObjectFlattener();
            bool hasExplicitColumns = options.Columns is { Length: > 0 };
            bool hasProjectionRules = hasExplicitColumns
                || options.Ignore?.Length > 0
                || options.IncludeProperties?.Length > 0
                || options.ExcludeProperties?.Length > 0
                || options.PinnedFirst?.Length > 0
                || options.PinnedLast?.Length > 0
                || options.PropertyPriority.Count > 0
                || options.MaxColumns < sourceColumnNames.Length;
            List<string> paths;
            if (!hasProjectionRules) {
                EnsureDataTableColumnLimit(sourceColumnNames.Length, configuredMaxColumns);
                paths = sourceColumnNames.ToList();
            } else if (hasExplicitColumns) {
                paths = ResolveExplicitDataTablePaths(flattener, options);
            } else {
                try {
                    paths = flattener.ResolvePaths(sourceColumnNames, options);
                } catch (InvalidDataException exception) {
                    if (exception.Data["OfficeIMO.RequiredColumns"] is int requiredColumns) {
                        EnsureDataTableColumnLimit(requiredColumns, configuredMaxColumns);
                    }
                    throw;
                }
            }
            if (paths.Count == 0) {
                Sheet.Cell(_row, 1, "(no tabular columns for the DataTable schema)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            int headerRow = _row;
            int lastRow = checked(headerRow + table.Rows.Count);
            if (lastRow > A1.MaxRows) {
                throw new InvalidDataException(
                    $"TableFrom(DataTable) requires rows {headerRow} through {lastRow}, exceeding Excel's {A1.MaxRows}-row worksheet limit. "
                    + "Split the data across multiple worksheets or start the table earlier; Excel's worksheet row limit cannot be overridden.");
            }

            List<string> headers = BuildTransformedHeaders(paths, options);
            EnsureUniqueTableHeaders(headers);
            var source = new DataTableTabularRowSource(table, paths, headers);
            string range = Sheet.InsertTabularRowSourceAsTableForDeferredMaterialization(
                source,
                startRow: headerRow,
                startColumn: 1,
                includeHeaders: true,
                tableName: title ?? "Table",
                style: style,
                includeAutoFilter: autoFilter);
            if (string.IsNullOrEmpty(range)) {
                throw new InvalidOperationException("The fixed-schema DataTable could not be inserted into the target worksheet.");
            }

            return CompleteTable(range, paths, headers, headerRow, lastRow, style, freezeHeaderRow, visuals);
        }

        private static List<string> ResolveExplicitDataTablePaths(
            ObjectFlattener flattener,
            ObjectFlattenerOptions options) {
            var paths = new List<string>();
            var added = new HashSet<string>(StringComparer.Ordinal);
            foreach (string candidate in options.Columns!) {
                if (string.IsNullOrWhiteSpace(candidate) || !added.Add(candidate)) continue;
                List<string> selected = flattener.ResolvePaths(new[] { candidate }, options);
                if (selected.Count == 0) continue;
                if (paths.Count >= options.MaxColumns) {
                    EnsureDataTableColumnLimit(checked(paths.Count + 1), options.MaxColumns);
                }
                paths.Add(selected[0]);
            }
            return paths;
        }

        private static void EnsureDataTableColumnLimit(int requiredColumns, int configuredMaxColumns) {
            if (requiredColumns <= configuredMaxColumns && requiredColumns <= A1.MaxColumns) return;
            if (requiredColumns > A1.MaxColumns && configuredMaxColumns >= A1.MaxColumns) {
                throw CreateExcelColumnLimitException(requiredColumns);
            }

            throw new InvalidDataException(
                $"TableFrom(DataTable) requires at least {requiredColumns} columns, exceeding the {configuredMaxColumns}-column materialization limit (MaxColumns). "
                + $"If this projection is intentional, raise the limit with configure: options => options.MaxColumns = {requiredColumns}.");
        }

        private static InvalidDataException CreateExcelColumnLimitException(int requiredColumns) =>
            new InvalidDataException(
                $"TableFrom(DataTable) requires at least {requiredColumns} columns, exceeding Excel's {A1.MaxColumns}-column worksheet limit. "
                + "Select fewer columns or split the data across multiple worksheets; Excel's worksheet column limit cannot be overridden.");
    }
}
