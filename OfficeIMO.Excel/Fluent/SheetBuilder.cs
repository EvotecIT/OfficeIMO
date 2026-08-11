using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Data;
using System.Reflection;
using System.Text;
using OfficeIMO.Data;
using OfficeIMO.Spreadsheet;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for composing a worksheet: headers, rows, ranges, tables, styles and filters.
    /// </summary>
    public class SheetBuilder {
        private static readonly ConcurrentDictionary<Type, RowsFromSimpleTypePlan> RowsFromSimpleTypePlans = new();
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
            Sheet = _fluent.Workbook.AddWorksheet(name);
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
        public SheetBuilder RowsFrom<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (data == null) throw new ArgumentNullException(nameof(data));

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            if (options.MaxRows <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxRows));
            options.MaxColumns = Math.Min(options.MaxColumns, A1.MaxColumns);
            int maximumDataRows = Math.Min(options.MaxRows, Math.Max(0, A1.MaxRows - _currentRow));
            int startRow = _currentRow;
            if (configure == null && !ObjectDictionaryAdapter.IsDictionaryType(typeof(T))) {
                RowsFromSimpleTypePlan typePlan = GetRowsFromSimpleTypePlan(typeof(T));
                if (typePlan.CanUseDirectSave) {
                    EnsureRowsFromColumnLimit(typePlan.Headers.Length, options.MaxColumns);
                    int maximumSimpleRows = GetRowsFromMaximumRowsForCellLimit(
                        maximumDataRows,
                        typePlan.Headers.Length,
                        options.MaxCells);
                    IReadOnlyList<T> simpleRows = MaterializeRowsBounded(
                        data,
                        maximumSimpleRows,
                        typePlan.Headers.Length,
                        options.MaxCells);
                    if (simpleRows.Count == 0) return this;
                    if (TryRowsFromSimpleFastPath(simpleRows, startRow, options)) return this;
                }
            }

            if (maximumDataRows == 0) {
                using IEnumerator<T> enumerator = data.GetEnumerator();
                if (enumerator.MoveNext()) {
                    throw new InvalidDataException("RowsFrom exceeds the 0-row worksheet boundary.");
                }
                return this;
            }
            options.MaxRows = maximumDataRows;
            var flattener = new ObjectFlattener();
            ObjectTableProjection projection = flattener.FlattenRows(data, options, "RowsFrom", headerRowCount: 1);
            if (projection.Rows.Count == 0) return this;
            IReadOnlyList<string> paths = projection.Columns;
            var headers = BuildTransformedHeaders(paths, options);

            var rowValues = new List<object?[]>(projection.Rows.Count);
            int dataRows = 0;
            foreach (Dictionary<string, object?> dict in projection.Rows) {
                if (options.CollectionMode == CollectionMode.ExpandRows) {
                    var collectionPath = paths.FirstOrDefault(p => dict.TryGetValue(p, out var val) && val is IEnumerable && val is not string);
                    if (collectionPath != null && dict[collectionPath] is IEnumerable coll) {
                        dataRows += AddExpandedRowsFromCollection(
                            rowValues,
                            paths,
                            dict,
                            options.DefaultValues,
                            collectionPath,
                            coll,
                            maximumDataRows - dataRows,
                            options.MaxCollectionItems,
                            options.MaxCells,
                            dataRows);
                        continue;
                    }
                }

                if (dataRows >= maximumDataRows) {
                    throw new InvalidDataException($"RowsFrom exceeds the {maximumDataRows}-row materialization limit.");
                }
                EnsureRowsFromCellLimit(dataRows + 1, paths.Count, options.MaxCells);
                rowValues.Add(ProjectRowsFromValues(paths, dict, options.DefaultValues));
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

        private static IReadOnlyList<T> MaterializeRowsBounded<T>(
            IEnumerable<T> data,
            int maximumRows,
            int cellColumnCount,
            long maximumCells) {
            if (data is IReadOnlyList<T> rows) {
                if (rows.Count > maximumRows) {
                    EnsureRowsFromCellLimit(rows.Count, cellColumnCount, maximumCells);
                    throw new InvalidDataException($"RowsFrom exceeds the {maximumRows}-row materialization limit.");
                }
                return rows;
            }

            var materialized = new List<T>(Math.Min(maximumRows, 4096));
            foreach (T item in data) {
                if (materialized.Count >= maximumRows) {
                    EnsureRowsFromCellLimit(materialized.Count + 1, cellColumnCount, maximumCells);
                    throw new InvalidDataException($"RowsFrom exceeds the {maximumRows}-row materialization limit.");
                }
                materialized.Add(item);
            }
            return materialized;
        }

        private static int GetRowsFromMaximumRowsForCellLimit(
            int maximumRows,
            int columnCount,
            long maximumCells) {
            if (columnCount <= 0) return maximumRows;
            long cellBoundedRows = Math.Max(0L, maximumCells / columnCount - 1L);
            return (int)Math.Min(maximumRows, cellBoundedRows);
        }

        private static int AddExpandedRowsFromCollection(
            List<object?[]> rowValues,
            IReadOnlyList<string> paths,
            Dictionary<string, object?> values,
            IReadOnlyDictionary<string, object?> defaultValues,
            string collectionPath,
            IEnumerable collection,
            int maximumRows,
            int maximumCollectionItems,
            long maximumCells,
            int existingRows) {
            int added = 0;
            foreach (object? element in collection) {
                if (added >= maximumCollectionItems) {
                    throw new InvalidDataException($"RowsFrom nested collection exceeds the {maximumCollectionItems}-item flattening limit.");
                }
                if (added >= maximumRows) {
                    throw new InvalidDataException($"RowsFrom nested expansion exceeds the configured row materialization limit.");
                }
                EnsureRowsFromCellLimit(existingRows + added + 1, paths.Count, maximumCells);
                rowValues.Add(ProjectRowsFromValues(paths, values, defaultValues, collectionPath, element));
                added++;
            }

            if (added == 0) {
                if (maximumRows <= 0) {
                    throw new InvalidDataException("RowsFrom nested expansion exceeds the configured row materialization limit.");
                }
                EnsureRowsFromCellLimit(existingRows + 1, paths.Count, maximumCells);
                rowValues.Add(ProjectRowsFromValues(paths, values, defaultValues: null));
                return 1;
            }

            return added;
        }

        private static object?[] ProjectRowsFromValues(
            IReadOnlyList<string> paths,
            Dictionary<string, object?> values,
            IReadOnlyDictionary<string, object?>? defaultValues,
            string? collectionPath = null,
            object? collectionValue = null) {
            var projected = new object?[paths.Count];
            for (int i = 0; i < paths.Count; i++) {
                string path = paths[i];
                if (collectionPath != null && string.Equals(path, collectionPath, StringComparison.Ordinal)) {
                    projected[i] = collectionValue;
                    continue;
                }

                if (values.TryGetValue(path, out object? value)) {
                    projected[i] = value;
                } else if (defaultValues != null && defaultValues.TryGetValue(path, out object? defaultValue)) {
                    projected[i] = defaultValue;
                }
            }

            return projected;
        }

        private bool TryRowsFromSimpleFastPath<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IReadOnlyList<T> rows, int startRow, ObjectFlattenerOptions options) {
            if (Sheet == null) {
                return false;
            }

            var typePlan = GetRowsFromSimpleTypePlan(typeof(T));
            if (!typePlan.CanUseDirectSave) {
                return false;
            }

            EnsureRowsFromColumnLimit(typePlan.Headers.Length, options.MaxColumns);
            EnsureRowsFromCellLimit(rows.Count, typePlan.Headers.Length, options.MaxCells);
            var directRows = MaterializeSimpleRowsFromProperties(typePlan, rows, out var columnTypes);
            int tableEndRow = startRow + rows.Count;
            string[] headers = typePlan.Headers;
            string range = $"A{startRow}:{ColumnLetter(headers.Length)}{tableEndRow}";
            if (startRow == 1 && Sheet.TryInsertRowsAsDeferredDirectSave("RowsFrom", headers, columnTypes, directRows, startRow, includeHeaders: true, range: range)) {
                _currentRow = tableEndRow + 1;
                _lastRange = range;
                return true;
            }

            AddSimpleRowsFromCellValues(startRow, headers, directRows);
            _currentRow = tableEndRow + 1;
            _lastRange = range;
            return true;
        }

        private static void EnsureRowsFromCellLimit(int dataRowCount, int columnCount, long maximumCells) {
            long projectedCells = ((long)dataRowCount + 1L) * columnCount;
            if (projectedCells > maximumCells) {
                throw new InvalidDataException(
                    $"RowsFrom exceeds the {maximumCells}-cell materialization limit.");
            }
        }

        internal static void EnsureRowsFromColumnLimit(int columnCount, int maximumColumns) {
            if (columnCount > maximumColumns) {
                throw new InvalidDataException(
                    $"RowsFrom exceeds the {maximumColumns}-column materialization limit.");
            }
        }

        /// <summary>Adds a table over the last added block (from RowsFrom) using the specified name.</summary>
        public SheetBuilder Table(string name, Action<ExcelTableBuilder>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (string.IsNullOrEmpty(_lastRange)) throw new InvalidOperationException("RowsFrom must be called before Table");
            var builder = new ExcelTableBuilder(Sheet);
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

            var cells = new List<(int Row, int Column, object Value)>(Math.Max(1, rowCount * colCount));
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

            SpreadsheetRangeReference syntax = SpreadsheetRangeReference.Parse(reference, SpreadsheetAddressDialect.ExcelA1);
            if (!syntax.Start.IsCell || !(syntax.End ?? syntax.Start).IsCell) {
                throw new ArgumentException("Range must contain cell endpoints.", nameof(reference));
            }
            if (syntax.Start.SheetName != null || syntax.End?.SheetName != null) {
                throw new ArgumentException("Range must not qualify a different worksheet.", nameof(reference));
            }
            SpreadsheetCellReference end = syntax.End ?? syntax.Start;
            int fromRow = checked((int)syntax.Start.Row!.Value);
            int fromCol = syntax.Start.Column!.Value;
            int toRow = checked((int)end.Row!.Value);
            int toCol = end.Column!.Value;

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
        public SheetBuilder Table(Action<ExcelTableBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new ExcelTableBuilder(Sheet);
            action(builder);
            // Note: The ExcelTableBuilder will handle AutoFilter conflicts internally
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
                HeaderCase.Pascal => TransformHeaderPascal(path),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

        private static List<string> BuildTransformedHeaders(IReadOnlyList<string> paths, ObjectFlattenerOptions options) {
            var headers = new List<string>(paths.Count);
            for (int i = 0; i < paths.Count; i++) {
                headers.Add(TransformHeader(paths[i], options));
            }

            return headers;
        }

        private static string TransformHeaderPascal(string path) {
            if (path.Length == 0) {
                ThrowEmptyHeaderSegment();
            }

            var builder = new StringBuilder(path.Length);
            int segmentStart = 0;
            for (int i = 0; i <= path.Length; i++) {
                if (i < path.Length && path[i] != '.') {
                    continue;
                }

                AppendPascalSegment(builder, path, segmentStart, i - segmentStart);
                segmentStart = i + 1;
            }

            return builder.ToString();
        }

        private static void AppendPascalSegment(StringBuilder builder, string path, int start, int length) {
            if (length == 0) {
                ThrowEmptyHeaderSegment();
            }

            builder.Append(char.ToUpperInvariant(path[start]));
            if (length > 1) {
                builder.Append(path, start + 1, length - 1);
            }
        }

        private static void ThrowEmptyHeaderSegment() => throw new IndexOutOfRangeException();

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

        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "RowsFrom limits inferred DataColumn types to supported scalar framework types; custom runtime types fall back to object/string cell storage.")]
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

        private static RowsFromSimpleTypePlan GetRowsFromSimpleTypePlan(Type type)
            => RowsFromSimpleTypePlans.GetOrAdd(type, CreateRowsFromSimpleTypePlan);

        [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "RowsFrom<T> preserves public properties on T; this cache only discovers those public row properties.")]
        private static RowsFromSimpleTypePlan CreateRowsFromSimpleTypePlan(Type type) {
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.GetIndexParameters().Length == 0 && property.GetMethod != null)
                .OrderBy(ObjectFlattener.GetMetadataOrder)
                .ThenBy(property => property.Name, StringComparer.Ordinal)
                .ToArray();
            if (properties.Length == 0 || properties.Any(property => !IsRowsFromDirectSaveScalarType(property.PropertyType))) {
                return RowsFromSimpleTypePlan.NotSupported;
            }

            var headers = new string[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
                headers[i] = properties[i].Name;
            }

            var staticColumnTypes = new Type[properties.Length];
            var staticBlankAsEmptyString = new bool[properties.Length];
            var staticRequiresBlankCoercion = new bool[properties.Length];
            var inferenceFallbackTypes = new Type[properties.Length];
            var inferColumns = new bool[properties.Length];
            var getters = new RowsFromSimpleValueGetter[properties.Length];
            var inferenceColumnIndexes = new List<int>();
            bool hasInferenceColumns = false;
            for (int i = 0; i < properties.Length; i++) {
                getters[i] = CreateRowsFromSimpleValueGetter(properties[i]);
                Type propertyType = properties[i].PropertyType;
                Type? nullableUnderlyingType = Nullable.GetUnderlyingType(propertyType);
                Type declaredType = nullableUnderlyingType ?? propertyType;
                if (declaredType == typeof(string)) {
                    staticColumnTypes[i] = typeof(string);
                    staticBlankAsEmptyString[i] = true;
                    staticRequiresBlankCoercion[i] = true;
                    inferenceFallbackTypes[i] = typeof(string);
                } else if (nullableUnderlyingType != null && CanUseStaticNullableRowsFromColumnType(nullableUnderlyingType)) {
                    staticColumnTypes[i] = nullableUnderlyingType;
                    staticBlankAsEmptyString[i] = false;
                    staticRequiresBlankCoercion[i] = true;
                    inferenceFallbackTypes[i] = nullableUnderlyingType;
                } else if (!declaredType.IsValueType || nullableUnderlyingType != null) {
                    inferColumns[i] = true;
                    inferenceColumnIndexes.Add(i);
                    hasInferenceColumns = true;
                    inferenceFallbackTypes[i] = typeof(object);
                } else {
                    staticColumnTypes[i] = declaredType;
                    staticBlankAsEmptyString[i] = false;
                    staticRequiresBlankCoercion[i] = false;
                    inferenceFallbackTypes[i] = declaredType;
                }
            }

            return CanUseRowsFromDataTable(headers)
                ? new RowsFromSimpleTypePlan(getters, headers, staticColumnTypes, staticBlankAsEmptyString, staticRequiresBlankCoercion, inferenceFallbackTypes, inferColumns, inferenceColumnIndexes.ToArray(), hasInferenceColumns, canUseDirectSave: true)
                : RowsFromSimpleTypePlan.NotSupported;
        }

        private static RowsFromSimpleValueGetter CreateRowsFromSimpleValueGetter(PropertyInfo property) {
            return property.GetValue;
        }

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

        private static bool CanUseStaticNullableRowsFromColumnType(Type type) {
            if (type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)) {
                return false;
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly) || type == typeof(TimeOnly)) {
                return false;
            }
#endif

            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(decimal)
                || type == typeof(Guid);
        }

        private static object?[][] MaterializeSimpleRowsFromProperties<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, out Type[] columnTypes) {
            var directRows = new object?[rows.Count][];
            if (!typePlan.HasInferenceColumns) {
                MaterializeStaticSimpleRowsFromProperties(typePlan, rows, directRows);
                columnTypes = typePlan.StaticColumnTypes;
                return directRows;
            }

            columnTypes = InferRowsFromPropertyColumnTypes(typePlan, rows, directRows);
            int[] inferenceColumnIndexes = typePlan.InferenceColumnIndexes;
            for (int row = 0; row < directRows.Length; row++) {
                object?[] values = directRows[row];
                for (int i = 0; i < inferenceColumnIndexes.Length; i++) {
                    int column = inferenceColumnIndexes[i];
                    values[column] = CoerceRowsFromDataTableValue(values[column], columnTypes[column]);
                }
            }

            return directRows;
        }

        private static void MaterializeStaticSimpleRowsFromProperties<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, object?[][] directRows) {
            RowsFromSimpleValueGetter[] getters = typePlan.Getters;
            bool[] blankAsEmptyString = typePlan.StaticBlankAsEmptyString;
            bool[] requiresBlankCoercion = typePlan.StaticRequiresBlankCoercion;
            int columnCount = getters.Length;
            for (int row = 0; row < rows.Count; row++) {
                T sourceRow = rows[row];
                var values = new object?[columnCount];
                for (int column = 0; column < columnCount; column++) {
                    object? value = getters[column](sourceRow);
                    values[column] = requiresBlankCoercion[column]
                        ? CoerceRowsFromDirectSaveValue(value, blankAsEmptyString[column])
                        : value;
                }

                directRows[row] = values;
            }
        }

        private static Type[] InferRowsFromPropertyColumnTypes<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, object?[][] directRows) {
            RowsFromSimpleValueGetter[] getters = typePlan.Getters;
            bool[] inferColumns = typePlan.InferColumns;
            bool[] staticBlankAsEmptyString = typePlan.StaticBlankAsEmptyString;
            bool[] staticRequiresBlankCoercion = typePlan.StaticRequiresBlankCoercion;
            int columnCount = getters.Length;
            var columnTypes = new Type[columnCount];
            Array.Copy(typePlan.StaticColumnTypes, columnTypes, columnTypes.Length);
            Type?[]? inferredTypes = typePlan.HasInferenceColumns ? new Type?[columnCount] : null;

            for (int row = 0; row < rows.Count; row++) {
                T sourceRow = rows[row];
                var values = new object?[columnCount];
                for (int column = 0; column < columnCount; column++) {
                    object? value = getters[column](sourceRow);
                    if (!inferColumns[column]) {
                        values[column] = staticRequiresBlankCoercion[column]
                            ? CoerceRowsFromDirectSaveValue(value, staticBlankAsEmptyString[column])
                            : value;
                        continue;
                    }

                    values[column] = value;
                    if (IsRowsFromBlankValue(value)) {
                        continue;
                    }

                    Type valueType = value!.GetType();
                    Type? inferred = inferredTypes![column];
                    if (inferred == null) {
                        inferredTypes[column] = valueType;
                    } else if (inferred != valueType) {
                        inferredTypes[column] = typeof(object);
                    }
                }

                directRows[row] = values;
            }

            int[] inferenceColumnIndexes = typePlan.InferenceColumnIndexes;
            for (int i = 0; i < inferenceColumnIndexes.Length; i++) {
                int column = inferenceColumnIndexes[i];
                columnTypes[column] = inferredTypes![column] ?? typePlan.InferenceFallbackTypes[column];
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

                    Type valueType = value!.GetType();
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

        private static object CoerceRowsFromDirectSaveValue(object? value, bool blankAsEmptyString) {
            if (IsRowsFromBlankValue(value)) {
                return blankAsEmptyString ? string.Empty : DBNull.Value;
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

        private void AddSimpleRowsFromCellValues(int startRow, IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            int totalCellCount = checked((rowValues.Count + 1) * headers.Count);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            for (int i = 0; i < headers.Count; i++) {
                cells[cellIndex++] = (startRow, i + 1, headers[i]);
            }

            for (int r = 0; r < rowValues.Count; r++) {
                object?[] values = rowValues[r];
                for (int c = 0; c < headers.Count; c++) {
                    cells[cellIndex++] = (startRow + r + 1, c + 1, values[c] ?? string.Empty);
                }
            }

            Sheet!.CellValues(cells);
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
            SpreadsheetRangeReference syntax = SpreadsheetRangeReference.Parse(reference, SpreadsheetAddressDialect.ExcelA1);
            if (syntax.End != null || !syntax.Start.IsCell || syntax.Start.SheetName != null) {
                throw new ArgumentException("Cell reference must identify one cell on the current worksheet.", nameof(reference));
            }
            return (checked((int)syntax.Start.Row!.Value), syntax.Start.Column!.Value);
        }

        private sealed class RowsFromSimpleTypePlan {
            internal static readonly RowsFromSimpleTypePlan NotSupported = new(
                Array.Empty<RowsFromSimpleValueGetter>(),
                Array.Empty<string>(),
                Array.Empty<Type>(),
                Array.Empty<bool>(),
                Array.Empty<bool>(),
                Array.Empty<Type>(),
                Array.Empty<bool>(),
                Array.Empty<int>(),
                hasInferenceColumns: false,
                canUseDirectSave: false);

            internal RowsFromSimpleTypePlan(
                RowsFromSimpleValueGetter[] getters,
                string[] headers,
                Type[] staticColumnTypes,
                bool[] staticBlankAsEmptyString,
                bool[] staticRequiresBlankCoercion,
                Type[] inferenceFallbackTypes,
                bool[] inferColumns,
                int[] inferenceColumnIndexes,
                bool hasInferenceColumns,
                bool canUseDirectSave) {
                Getters = getters;
                Headers = headers;
                StaticColumnTypes = staticColumnTypes;
                StaticBlankAsEmptyString = staticBlankAsEmptyString;
                StaticRequiresBlankCoercion = staticRequiresBlankCoercion;
                InferenceFallbackTypes = inferenceFallbackTypes;
                InferColumns = inferColumns;
                InferenceColumnIndexes = inferenceColumnIndexes;
                HasInferenceColumns = hasInferenceColumns;
                CanUseDirectSave = canUseDirectSave;
            }

            internal RowsFromSimpleValueGetter[] Getters { get; }

            internal string[] Headers { get; }

            internal Type[] StaticColumnTypes { get; }

            internal bool[] StaticBlankAsEmptyString { get; }

            internal bool[] StaticRequiresBlankCoercion { get; }

            internal Type[] InferenceFallbackTypes { get; }

            internal bool[] InferColumns { get; }

            internal int[] InferenceColumnIndexes { get; }

            internal bool HasInferenceColumns { get; }

            internal bool CanUseDirectSave { get; }
        }

        private delegate object? RowsFromSimpleValueGetter(object? row);
    }
}
