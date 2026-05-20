using System.Globalization;
using System.Data;
using System.Reflection;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for composing a worksheet: headers, rows, ranges, tables, styles and filters.
    /// </summary>
    public class SheetBuilder {
        private readonly ExcelFluentWorkbook _fluent;
        internal ExcelSheet? Sheet { get; private set; }
        private int _currentRow = 1;
        private string? _lastRange;
        private string? _pendingAutoFilterRange;
        private Dictionary<uint, IEnumerable<string>>? _pendingAutoFilterCriteria;

        internal SheetBuilder(ExcelFluentWorkbook fluent) {
            _fluent = fluent;
        }

        /// <summary>Creates and selects a new worksheet.</summary>
        /// <param name="name">Optional sheet name.</param>
        public SheetBuilder AddSheet(string name = "") {
            Sheet = _fluent.Workbook.AddWorkSheet(name);
            return this;
        }

        /// <summary>Adds a header row with the provided values.</summary>
        public SheetBuilder HeaderRow(params object?[] values) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var row = new RowBuilder(this, Sheet, _currentRow);
            row.Values(values);
            _currentRow++;
            return this;
        }

        /// <summary>Adds a data row configured by the supplied builder action.</summary>
        public SheetBuilder Row(Action<RowBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new RowBuilder(this, Sheet, _currentRow);
            action(builder);
            _currentRow++;
            return this;
        }

        /// <summary>
        /// Generates rows from a sequence of objects using the object flattener.
        /// Produces a header row from flattened property paths, then data rows.
        /// </summary>
        public SheetBuilder RowsFrom<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (data == null) throw new ArgumentNullException(nameof(data));

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            var flattener = new ObjectFlattener();

            var enumerable = data.ToList();
            if (!enumerable.Any()) return this;

            int startRow = _currentRow;
            if (configure == null && TryRowsFromSimpleDirectSave(enumerable, startRow)) {
                return this;
            }

            var paths = options.Columns?.ToList() ?? flattener.GetPaths(typeof(T), options);
            var headers = paths.Select(p => TransformHeader(p, options)).ToList();

            var rowValues = new List<object?[]>();
            int dataRows = 0;
            foreach (var item in enumerable) {
                var dict = flattener.Flatten(item, options);
                if (options.CollectionMode == CollectionMode.ExpandRows) {
                    var collectionPath = paths.FirstOrDefault(p => dict.TryGetValue(p, out var val) && val is IEnumerable && val is not string);
                    if (collectionPath != null && dict[collectionPath] is IEnumerable coll) {
                        var list = coll.Cast<object?>().ToList();
                        if (list.Count == 0) {
                            rowValues.Add(paths.Select(p => dict.TryGetValue(p, out var v) ? v : null).ToArray());
                            dataRows++;
                        } else {
                            foreach (var element in list) {
                                var values = paths.Select(p => p == collectionPath ? element : dict.TryGetValue(p, out var v) ? v : (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray();
                                rowValues.Add(values);
                                dataRows++;
                            }
                        }
                        continue;
                    }
                }

                rowValues.Add(paths.Select(p => dict.TryGetValue(p, out var v) ? v : (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray());
                dataRows++;
            }

            if (CanUseRowsFromDataTable(headers)) {
                var table = CreateRowsFromDataTable(headers, rowValues);
                int tableEndRow = startRow + dataRows;
                string range = $"A{startRow}:{ColumnLetter(headers.Count)}{tableEndRow}";
                if (!Sheet.TryInsertOwnedDataTableAsDeferredDirectSave(table, startRow, includeHeaders: true, range: range)) {
                    Sheet.InsertOwnedDataTable(table, startRow, startColumn: 1, includeHeaders: true);
                }
            } else {
                var cells = new List<(int Row, int Column, object Value)>((dataRows + 1) * Math.Max(1, headers.Count));
                AddRowsFromCellValues(cells, startRow, headers, rowValues);
                Sheet.CellValues(cells);
            }

            _currentRow = startRow + dataRows + 1;
            int endRow = _currentRow - 1;
            _lastRange = $"A{startRow}:{ColumnLetter(headers.Count)}{endRow}";

            return this;
        }

        private bool TryRowsFromSimpleDirectSave<T>(IReadOnlyList<T> rows, int startRow) {
            if (Sheet == null || startRow != 1) {
                return false;
            }

            var properties = GetSimpleRowsFromProperties(typeof(T));
            if (properties.Length == 0) {
                return false;
            }

            if (properties.Any(property => !IsRowsFromDirectSaveScalarType(property.PropertyType))) {
                return false;
            }

            var headers = properties.Select(property => property.Name).ToArray();
            if (!CanUseRowsFromDataTable(headers)) {
                return false;
            }

            var columnTypes = InferRowsFromPropertyColumnTypes(properties, rows);
            var directRows = CreateRowsFromSimpleRows(properties, columnTypes, rows);
            int tableEndRow = startRow + rows.Count;
            string range = $"A{startRow}:{ColumnLetter(headers.Length)}{tableEndRow}";
            if (!Sheet.TryInsertRowsAsDeferredDirectSave("RowsFrom", headers, columnTypes, directRows, startRow, includeHeaders: true, range: range)) {
                return false;
            }

            _currentRow = tableEndRow + 1;
            _lastRange = range;
            return true;
        }

        /// <summary>Adds a table over the last added block (from RowsFrom) using the specified name.</summary>
        public SheetBuilder Table(string name, Action<TableBuilder>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (string.IsNullOrEmpty(_lastRange)) throw new InvalidOperationException("RowsFrom must be called before Table");
            var builder = new TableBuilder(Sheet);
            configure?.Invoke(builder);
            builder.Build(_lastRange!, name);
            return this;
        }

        /// <summary>Writes a cell at the specified row/column with optional value, formula and number format.</summary>
        public SheetBuilder Cell(int row, int column, object? value = null, string? formula = null, string? numberFormat = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
            Sheet.Cell(row, column, value, formula, numberFormat);
            return this;
        }

        /// <summary>Writes a cell using A1 reference with optional value, formula and number format.</summary>
        public SheetBuilder Cell(string reference, object? value = null, string? formula = null, string? numberFormat = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (string.IsNullOrWhiteSpace(reference)) throw new ArgumentNullException(nameof(reference));
            var (row, column) = ParseCellReference(reference);
            Sheet.Cell(row, column, value, formula, numberFormat);
            return this;
        }

        /// <summary>
        /// Writes a rectangular block specified by coordinates, optionally providing a 2D values array.
        /// </summary>
        public SheetBuilder Range(int fromRow, int fromCol, int toRow, int toCol, object[,]? values = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (fromRow < 1) throw new ArgumentOutOfRangeException(nameof(fromRow));
            if (fromCol < 1) throw new ArgumentOutOfRangeException(nameof(fromCol));
            if (toRow < 1) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toCol < 1) throw new ArgumentOutOfRangeException(nameof(toCol));
            if (toRow < fromRow) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toCol < fromCol) throw new ArgumentOutOfRangeException(nameof(toCol));

            int rowCount = toRow - fromRow + 1;
            int colCount = toCol - fromCol + 1;

            if (values != null && (values.GetLength(0) != rowCount || values.GetLength(1) != colCount)) {
                throw new ArgumentException("Values array dimensions must match the specified range.", nameof(values));
            }

            var cells = new List<(int Row, int Column, object Value)>();
            for (int r = 0; r < rowCount; r++) {
                for (int c = 0; c < colCount; c++) {
                    object cellValue = values != null ? values[r, c] : string.Empty;
                    cells.Add((fromRow + r, fromCol + c, cellValue));
                }
            }

            Sheet.CellValues(cells);
            return this;
        }

        /// <summary>
        /// Configures a range using A1 notation via the range builder (styles, merges, borders, etc.).
        /// </summary>
        public SheetBuilder Range(string reference, Action<RangeBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (string.IsNullOrWhiteSpace(reference)) throw new ArgumentNullException(nameof(reference));

            var parts = reference.Split(':');
            var start = parts[0];
            var end = parts.Length > 1 ? parts[1] : parts[0];

            var (fromRow, fromCol) = ParseCellReference(start);
            var (toRow, toCol) = ParseCellReference(end);

            var builder = new RangeBuilder(Sheet, fromRow, fromCol, toRow, toCol);
            action(builder);
            return this;
        }

        /// <summary>Configures column widths and styles via the column collection builder.</summary>
        public SheetBuilder Columns(Action<ColumnCollectionBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new ColumnCollectionBuilder(Sheet);
            action(builder);
            return this;
        }

        /// <summary>Adds and configures a table via the table builder.</summary>
        public SheetBuilder Table(Action<TableBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new TableBuilder(Sheet);
            action(builder);
            // Note: The TableBuilder will handle AutoFilter conflicts internally
            return this;
        }

        /// <summary>Applies ad‑hoc styles via the style builder.</summary>
        public SheetBuilder Style(Action<StyleBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new StyleBuilder(Sheet);
            action(builder);
            return this;
        }

        /// <summary>Applies AutoFilter to a range with optional per‑column criteria.</summary>
        public SheetBuilder AutoFilter(string range, Dictionary<uint, IEnumerable<string>>? criteria = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");

            // Store the pending AutoFilter for conflict detection
            _pendingAutoFilterRange = range;
            _pendingAutoFilterCriteria = criteria;

            // Apply the AutoFilter
            Sheet.AddAutoFilter(range, criteria);
            return this;
        }

        /// <summary>Adds a 2‑color scale conditional formatting rule over the range.</summary>
        public SheetBuilder ConditionalColorScale(string range, OfficeColor startColor, OfficeColor endColor) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.AddConditionalColorScale(range, startColor, endColor);
            return this;
        }

        /// <summary>Adds a data bar conditional formatting rule over the range.</summary>
        public SheetBuilder ConditionalDataBar(string range, OfficeColor color) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.AddConditionalDataBar(range, color);
            return this;
        }

        /// <summary>
        /// Freezes the specified number of rows and columns on the current sheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        public SheetBuilder Freeze(int topRows = 0, int leftCols = 0) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.Freeze(topRows, leftCols);
            return this;
        }

        /// <summary>Auto‑fits columns and/or rows in the current sheet.</summary>
        public SheetBuilder AutoFit(bool columns, bool rows) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (columns) {
                Sheet.AutoFitColumns();
            }
            if (rows) {
                Sheet.AutoFitRows();
            }
            return this;
        }

        private static string TransformHeader(string path, ObjectFlattenerOptions opts) {
            foreach (var prefix in opts.HeaderPrefixTrimPaths) {
                if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    path = path.Substring(prefix.Length);
                }
            }
            return opts.HeaderCase switch {
                HeaderCase.Pascal => string.Concat(path.Split('.').Select(s => char.ToUpperInvariant(s[0]) + s.Substring(1))),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

        private static bool CanUseRowsFromDataTable(IReadOnlyList<string> headers) {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string header in headers) {
                if (string.IsNullOrWhiteSpace(header)) {
                    return false;
                }

                if (!seen.Add(header)) {
                    return false;
                }
            }

            return headers.Count > 0;
        }

        private static DataTable CreateRowsFromDataTable(IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            var table = new DataTable("RowsFrom") {
                Locale = CultureInfo.InvariantCulture
            };

            var columnTypes = InferRowsFromColumnTypes(rowValues, headers.Count);
            for (int i = 0; i < headers.Count; i++) {
                table.Columns.Add(headers[i], columnTypes[i]);
            }

            table.BeginLoadData();
            try {
                foreach (object?[] values in rowValues) {
                    var row = new object[headers.Count];
                    for (int i = 0; i < headers.Count; i++) {
                        row[i] = CoerceRowsFromDataTableValue(i < values.Length ? values[i] : null, columnTypes[i]);
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static object?[][] CreateRowsFromSimpleRows<T>(IReadOnlyList<PropertyInfo> properties, IReadOnlyList<Type> columnTypes, IReadOnlyList<T> rows) {
            var directRows = new object?[rows.Count][];
            for (int r = 0; r < rows.Count; r++) {
                var row = new object?[properties.Count];
                for (int c = 0; c < properties.Count; c++) {
                    row[c] = CoerceRowsFromDataTableValue(properties[c].GetValue(rows[r]), columnTypes[c]);
                }

                directRows[r] = row;
            }

            return directRows;
        }

        private static PropertyInfo[] GetSimpleRowsFromProperties(Type type)
            => type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.GetIndexParameters().Length == 0 && property.GetMethod != null)
                .OrderBy(property => property.MetadataToken)
                .ToArray();

        private static bool IsRowsFromDirectSaveScalarType(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            if (type == typeof(string)) {
                return true;
            }

            if (typeof(System.Collections.IEnumerable).IsAssignableFrom(type)) {
                return false;
            }

            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid);
        }

        private static Type[] InferRowsFromPropertyColumnTypes<T>(IReadOnlyList<PropertyInfo> properties, IReadOnlyList<T> rows) {
            var columnTypes = new Type[properties.Count];
            for (int column = 0; column < properties.Count; column++) {
                Type declaredType = Nullable.GetUnderlyingType(properties[column].PropertyType) ?? properties[column].PropertyType;
                if (!declaredType.IsValueType || Nullable.GetUnderlyingType(properties[column].PropertyType) != null) {
                    Type? inferred = null;
                    for (int row = 0; row < rows.Count; row++) {
                        object? value = properties[column].GetValue(rows[row]);
                        if (IsRowsFromBlankValue(value)) {
                            continue;
                        }

                        Type valueType = Nullable.GetUnderlyingType(value!.GetType()) ?? value.GetType();
                        if (inferred == null) {
                            inferred = valueType;
                            continue;
                        }

                        if (inferred != valueType) {
                            inferred = typeof(object);
                            break;
                        }
                    }

                    columnTypes[column] = inferred ?? (declaredType == typeof(string) ? typeof(string) : typeof(object));
                } else {
                    columnTypes[column] = declaredType;
                }
            }

            return columnTypes;
        }

        private static Type[] InferRowsFromColumnTypes(IReadOnlyList<object?[]> rowValues, int columnCount) {
            var columnTypes = new Type[columnCount];
            for (int column = 0; column < columnCount; column++) {
                Type? inferred = null;
                for (int row = 0; row < rowValues.Count; row++) {
                    object? value = column < rowValues[row].Length ? rowValues[row][column] : null;
                    if (IsRowsFromBlankValue(value)) {
                        continue;
                    }

                    Type valueType = Nullable.GetUnderlyingType(value!.GetType()) ?? value.GetType();
                    if (inferred == null) {
                        inferred = valueType;
                        continue;
                    }

                    if (inferred != valueType) {
                        inferred = typeof(object);
                        break;
                    }
                }

                columnTypes[column] = inferred ?? typeof(string);
            }

            return columnTypes;
        }

        private static bool IsRowsFromBlankValue(object? value) {
            return value == null || value == DBNull.Value || value is string text && text.Length == 0;
        }

        private static object CoerceRowsFromDataTableValue(object? value, Type columnType) {
            if (IsRowsFromBlankValue(value)) {
                return columnType == typeof(string) || columnType == typeof(object)
                    ? string.Empty
                    : DBNull.Value;
            }

            return value!;
        }

        private static void AddRowsFromCellValues(List<(int Row, int Column, object Value)> cells, int startRow, IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            for (int i = 0; i < headers.Count; i++) {
                cells.Add((startRow, i + 1, headers[i]));
            }

            for (int r = 0; r < rowValues.Count; r++) {
                object?[] values = rowValues[r];
                for (int c = 0; c < headers.Count; c++) {
                    object? value = c < values.Length ? values[c] : null;
                    cells.Add((startRow + r + 1, c + 1, value ?? string.Empty));
                }
            }
        }

        private static string ColumnLetter(int column) {
            var dividend = column;
            var columnName = string.Empty;
            while (dividend > 0) {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static (int Row, int Column) ParseCellReference(string reference) {
            int i = 0;
            int col = 0;
            while (i < reference.Length && char.IsLetter(reference[i])) {
                col = col * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
                i++;
            }
            if (col == 0 || i >= reference.Length) throw new ArgumentException("Invalid cell reference", nameof(reference));
            var rowPart = reference.Substring(i);
            if (!int.TryParse(rowPart, out int row) || row <= 0) {
                throw new ArgumentException("Invalid cell reference", nameof(reference));
            }
            return (row, col);
        }
    }
}
