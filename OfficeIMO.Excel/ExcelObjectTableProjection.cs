using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Data;
using System.Reflection;
using System.Text;
using OfficeIMO.Data;
using OfficeIMO.Spreadsheet;
using OfficeColor = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel;

/// <summary>Preserves bounded flattening and deferred-cell projection for native object tables.</summary>
internal sealed partial class ExcelObjectTableProjection {
    private static readonly ConcurrentDictionary<Type, RowsFromSimpleTypePlan> RowsFromSimpleTypePlans = new();
    private readonly ExcelSheet Sheet;
    private readonly bool _includeHeaders;
    private readonly int _startColumn;
    private int _currentRow;
    private string? _lastRange;
    internal IReadOnlyList<string> ColumnNames { get; private set; } = Array.Empty<string>();
    internal static bool HasTableProjection(OfficeTableOptions? options) => options != null &&
        (options.Include.Count != 0 || options.Exclude.Count != 0 || options.Order.Count != 0 ||
         options.HeaderRenames.Count != 0 || options.Formatters.Count != 0 || options.HeaderTransform != null);

    private int HeaderRows => _includeHeaders ? 1 : 0;

    internal ExcelObjectTableProjection(ExcelSheet sheet, int row, int column, bool includeHeaders) {
        Sheet = sheet;
        _currentRow = row;
        _startColumn = column;
        _includeHeaders = includeHeaders;
    }
    private string BuildRange(int startRow, int endRow, int columns) =>
        ExcelA1.CellReference(startRow, _startColumn) + ":" + ExcelA1.CellReference(endRow, _startColumn + columns - 1);
        internal string? WriteRows<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IEnumerable<T> data, ObjectFlattenerOptions? projectionOptions = null, OfficeTableOptions? tableOptions = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (data == null) throw new ArgumentNullException(nameof(data));

            var options = projectionOptions?.CreateProjectionSnapshot() ?? new ObjectFlattenerOptions();
            OfficeIMO.Core.Internal.ObjectDataHelpers.ConfigureTableProjection(options, tableOptions);
            if (options.MaxRows <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxRows));
            options.MaxColumns = Math.Min(options.MaxColumns, ExcelA1.MaxColumns - _startColumn + 1);
            int maximumDataRows = Math.Min(options.MaxRows, Math.Max(0, ExcelA1.MaxRows - _currentRow + 1 - HeaderRows));
            int startRow = _currentRow;
            if (projectionOptions == null && !HasTableProjection(tableOptions) && !ObjectDictionaryAdapter.IsDictionaryType(typeof(T))) {
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
                    if (simpleRows.Count == 0) return _lastRange;
                    if (TryRowsFromSimpleFastPath(simpleRows, startRow, options)) return _lastRange;
                    data = simpleRows;
                }
            }

            if (maximumDataRows == 0) {
                using IEnumerator<T> enumerator = data.GetEnumerator();
                if (enumerator.MoveNext()) {
                    throw new InvalidDataException("AddTable exceeds the 0-row worksheet boundary.");
                }
                return _lastRange;
            }
            options.MaxRows = maximumDataRows;
            var flattener = new ObjectFlattener();
            ObjectTableProjection projection = flattener.FlattenRows(
                data,
                options,
                "AddTable",
                headerRowCount: HeaderRows,
                enforceEmptyProjectionLimits: false);
            if (projection.Rows.Count == 0) return _lastRange;
            IReadOnlyList<string> paths = projection.Columns;
            var headers = BuildTransformedHeaders(paths, options);
            ColumnNames = headers;

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
                    throw new InvalidDataException($"AddTable exceeds the {maximumDataRows}-row materialization limit.");
                }
                EnsureRowsFromCellLimit(dataRows + 1, paths.Count, options.MaxCells);
                rowValues.Add(ProjectRowsFromValues(paths, dict, options.DefaultValues));
                dataRows++;
            }

            if (CanUseRowsFromDataTable(headers)) {
                var table = CreateRowsFromDataTable(headers, rowValues);
                int tableEndRow = startRow + dataRows + HeaderRows - 1;
                string range = BuildRange(startRow, tableEndRow, headers.Count);
                if (_startColumn != 1 || !Sheet.TryInsertOwnedDataTableAsDeferredDirectSave(table, startRow, includeHeaders: _includeHeaders, range: range)) {
                    Sheet.InsertOwnedDataTable(table, startRow, startColumn: _startColumn, includeHeaders: _includeHeaders);
                }
            } else {
                var cells = new List<(int Row, int Column, object Value)>((dataRows + HeaderRows) * Math.Max(1, headers.Count));
                AddRowsFromCellValues(cells, startRow, headers, rowValues);
                Sheet.CellValues(cells);
            }

            _currentRow = startRow + dataRows + HeaderRows;
            int endRow = _currentRow - 1;
            _lastRange = BuildRange(startRow, endRow, headers.Count);

            return _lastRange;
        }

        private IReadOnlyList<T> MaterializeRowsBounded<T>(
            IEnumerable<T> data,
            int maximumRows,
            int cellColumnCount,
            long maximumCells) {
            if (data is IReadOnlyList<T> rows) {
                if (rows.Count > maximumRows) {
                    EnsureRowsFromCellLimit(rows.Count, cellColumnCount, maximumCells);
                    throw new InvalidDataException($"AddTable exceeds the {maximumRows}-row materialization limit.");
                }
                return rows;
            }

            var materialized = new List<T>(Math.Min(maximumRows, 4096));
            foreach (T item in data) {
                if (materialized.Count >= maximumRows) {
                    EnsureRowsFromCellLimit(materialized.Count + 1, cellColumnCount, maximumCells);
                    throw new InvalidDataException($"AddTable exceeds the {maximumRows}-row materialization limit.");
                }
                materialized.Add(item);
            }
            return materialized;
        }

        private int GetRowsFromMaximumRowsForCellLimit(
            int maximumRows,
            int columnCount,
            long maximumCells) {
            if (columnCount <= 0) return maximumRows;
            long cellBoundedRows = Math.Max(0L, maximumCells / columnCount - HeaderRows);
            return (int)Math.Min(maximumRows, cellBoundedRows);
        }

        private int AddExpandedRowsFromCollection(
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
                    throw new InvalidDataException($"AddTable nested collection exceeds the {maximumCollectionItems}-item flattening limit.");
                }
                if (added >= maximumRows) {
                    throw new InvalidDataException($"AddTable nested expansion exceeds the configured row materialization limit.");
                }
                EnsureRowsFromCellLimit(existingRows + added + 1, paths.Count, maximumCells);
                rowValues.Add(ProjectRowsFromValues(paths, values, defaultValues, collectionPath, element));
                added++;
            }

            if (added == 0) {
                if (maximumRows <= 0) {
                    throw new InvalidDataException("AddTable nested expansion exceeds the configured row materialization limit.");
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
            int tableEndRow = startRow + rows.Count + HeaderRows - 1;
            string[] headers = typePlan.Headers;
            ColumnNames = headers;
            string range = BuildRange(startRow, tableEndRow, headers.Length);
            if (_startColumn == 1 && startRow == 1 && Sheet.TryInsertRowsAsDeferredDirectSave("AddTable", headers, columnTypes, directRows, startRow, includeHeaders: _includeHeaders, range: range)) {
                _currentRow = tableEndRow + 1;
                _lastRange = range;
                return true;
            }

            AddSimpleRowsFromCellValues(startRow, headers, directRows);
            _currentRow = tableEndRow + 1;
            _lastRange = range;
            return true;
        }

        private void EnsureRowsFromCellLimit(int dataRowCount, int columnCount, long maximumCells) {
            long projectedCells = ((long)dataRowCount + HeaderRows) * columnCount;
            if (projectedCells > maximumCells) {
                throw new InvalidDataException(
                    $"AddTable exceeds the {maximumCells}-cell materialization limit.");
            }
        }

        internal static void EnsureRowsFromColumnLimit(int columnCount, int maximumColumns) {
            if (columnCount > maximumColumns) {
                throw new InvalidDataException(
                    $"AddTable exceeds the {maximumColumns}-column materialization limit.");
            }
        }


}
