using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int BufferedPowerShellObjectExportInitialColumnCapacity = 16;
        private const int BufferedPowerShellObjectExportColumnLimit = 64;
        private static readonly ConcurrentDictionary<Type, SimpleObjectExportPlan> SimpleObjectExportPlans = new();
        private static readonly ConcurrentDictionary<Type, PowerShellObjectExportPlan> PowerShellObjectExportPlans = new();

        /// <summary>
        /// Inserts objects into the worksheet by flattening their properties into columns.
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="includeHeaders">Whether to include column headers.</param>
        /// <param name="startRow">1-based starting row.</param>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly with CellValues or pre-flatten using known types.")]
        public void InsertObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(IEnumerable<T> items, bool includeHeaders = true, int startRow = 1) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var rows = items as IReadOnlyList<T> ?? items.ToList();
            if (rows.Count == 0) {
                return;
            }

            if (TryInsertSimpleObjectRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return;
            }

            if (TryInsertSimpleObjectRowsAsCellValues(rows, includeHeaders, startRow)) {
                return;
            }

            if (TryInsertFlatDictionaryRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return;
            }

            var flattenedItems = new List<Dictionary<string, object?>>(rows.Count);
            List<string> headers = new List<string>();
            HashSet<string> headerSet = new HashSet<string>();

            foreach (var item in rows) {
                var dict = new Dictionary<string, object?>();
                FlattenObject(item, null, dict);
                flattenedItems.Add(dict);
                foreach (var key in dict.Keys) {
                    if (headerSet.Add(key)) {
                        headers.Add(key);
                    }
                }
            }

            string? directSaveRange = null;
            bool hasBlankDisplayHeader = includeHeaders && headers.Any(string.IsNullOrWhiteSpace);
            if (!hasBlankDisplayHeader && CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                try {
                    directSaveRange = BuildObjectExportRange(startRow, headers.Count, flattenedItems.Count, includeHeaders);
                    var directRows = CreateObjectExportRows(headers, flattenedItems, out var columnTypes);
                    if (TryInsertRowsAsDeferredDirectSave(Name, headers, columnTypes, directRows, startRow, includeHeaders, directSaveRange)) {
                        return;
                    }
                } catch {
                    // Direct-save registration is opportunistic; fall back to the normal cell path.
                }
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * Math.Max(1, headers.Count));
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Count; c++) {
                    cells[cellIndex++] = (row, c + 1, headers[c]);
                }
                row++;
            }

            foreach (var dict in flattenedItems) {
                for (int c = 0; c < headers.Count; c++) {
                    object value = dict.TryGetValue(headers[c], out var entry) ? entry ?? string.Empty : string.Empty;
                    cells[cellIndex++] = (row, c + 1, value);
                }
                row++;
            }

            // Use the batch CellValues path with planner + execution policy
            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
        }

        /// <summary>
        /// Inserts objects into the worksheet using explicit column selectors (AOT-safe).
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="columns">Column headers and selectors.</param>
        public void InsertObjects<T>(IEnumerable<T> items, params (string Header, Func<T, object?> Selector)[] columns) {
            InsertObjects(items, includeHeaders: true, startRow: 1, columns);
        }

        /// <summary>
        /// Inserts objects into the worksheet using explicit column selectors (AOT-safe).
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="includeHeaders">Whether to include column headers.</param>
        /// <param name="startRow">1-based starting row.</param>
        /// <param name="columns">Column headers and selectors.</param>
        public void InsertObjects<T>(IEnumerable<T> items, bool includeHeaders, int startRow, params (string Header, Func<T, object?> Selector)[] columns) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }
            if (columns == null || columns.Length == 0) {
                throw new ArgumentException("At least one column selector is required.", nameof(columns));
            }

            var rows = items as IReadOnlyList<T> ?? items.ToList();
            if (rows.Count == 0) {
                return;
            }

            var headers = new string[columns.Length];
            var selectors = new Func<T, object?>[columns.Length];
            bool hasBlankDisplayHeader = false;
            for (int c = 0; c < columns.Length; c++) {
                string header = columns[c].Header ?? $"Column{c + 1}";
                headers[c] = header;
                selectors[c] = columns[c].Selector ?? NullObjectSelector;
                if (includeHeaders && string.IsNullOrWhiteSpace(header)) {
                    hasBlankDisplayHeader = true;
                }
            }

            object?[]? values = null;
            if (!hasBlankDisplayHeader
                && !HasDuplicateObjectExportHeaders(headers)
                && CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Length)) {
                Type?[] inferredColumnTypes;
                values = CreateExplicitObjectExportCellValues(rows, selectors, out inferredColumnTypes);
                try {
                    var columnTypes = CompleteObjectExportColumnTypes(inferredColumnTypes);
                    string directSaveRange = BuildObjectExportRange(startRow, headers.Length, rows.Count, includeHeaders);
                    if (TryInsertCellValuesAsDeferredDirectSave(Name, headers, columnTypes, values, headers.Length, rows.Count, startRow, includeHeaders, directSaveRange)) {
                        return;
                    }
                } catch {
                    // Direct-save registration is opportunistic; fall back to the normal cell path.
                }
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * headers.Length);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, headers[c]);
                }
                row++;
            }

            if (values != null) {
                for (int r = 0; r < rows.Count; r++) {
                    int rowOffset = r * headers.Length;
                    for (int c = 0; c < headers.Length; c++) {
                        cells[cellIndex++] = (row, c + 1, values[rowOffset + c] ?? string.Empty);
                    }

                    row++;
                }
            } else {
                foreach (var item in rows) {
                    for (int c = 0; c < headers.Length; c++) {
                        cells[cellIndex++] = (row, c + 1, selectors[c](item) ?? string.Empty);
                    }

                    row++;
                }
            }

            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
        }

        private static object? NullObjectSelector<T>(T row) => null;

        internal static object?[] CreateExplicitObjectExportCellValues<T>(
            IReadOnlyList<T> rows,
            IReadOnlyList<Func<T, object?>> selectors,
            out Type?[] inferredColumnTypes) {
            var values = new object?[checked(rows.Count * selectors.Count)];
            inferredColumnTypes = new Type?[selectors.Count];
            for (int r = 0; r < rows.Count; r++) {
                int rowOffset = r * selectors.Count;
                for (int c = 0; c < selectors.Count; c++) {
                    object? value = selectors[c](rows[r]);
                    values[rowOffset + c] = value;
                    UpdateObjectExportColumnType(inferredColumnTypes, c, value);
                }
            }

            return values;
        }
    }
}
