using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace OfficeIMO.Data {
    /// <summary>
    /// Provides the shared runtime-row projection used by Office table renderers.
    /// </summary>
    public partial class ObjectFlattener {
        internal ObjectTableProjection FlattenRows<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
            IEnumerable<T> source,
            ObjectFlattenerOptions options,
            string consumerName,
            int headerRowCount) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (headerRowCount < 0) throw new ArgumentOutOfRangeException(nameof(headerRowCount));
            ValidateLimits(options);

            List<T> items = MaterializeRowsBounded(source, options, consumerName);
            var rows = new List<Dictionary<string, object?>>(items.Count);
            var discoveredColumns = new List<string>();
            var discoveredColumnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (items.Count == 0
                && (options.Columns == null || options.Columns.Length == 0)
                && !ObjectDictionaryAdapter.IsDictionaryType(typeof(T))) {
                discoveredColumns.AddRange(GetPaths(typeof(T), options));
            }

            foreach (T item in items) {
                Dictionary<string, object?> row = Flatten(item, options);
                rows.Add(row);
                foreach (string path in row.Keys) {
                    if (!string.IsNullOrWhiteSpace(path) && discoveredColumnSet.Add(path)) {
                        if (discoveredColumns.Count >= options.MaxColumns) {
                            throw new InvalidDataException(
                                $"{consumerName} exceeds the {options.MaxColumns}-column materialization limit.");
                        }
                        discoveredColumns.Add(path);
                    }
                }
                EnsureTableCellLimit(rows.Count, discoveredColumns.Count, headerRowCount, options, consumerName);
            }

            IEnumerable<string> columnCandidates = options.Columns != null && options.Columns.Length > 0
                ? options.Columns
                : discoveredColumns;
            List<string> columns = ResolvePaths(columnCandidates, options);
            if (columns.Count > options.MaxColumns) {
                throw new InvalidDataException(
                    $"{consumerName} exceeds the {options.MaxColumns}-column materialization limit.");
            }
            EnsureTableCellLimit(rows.Count, columns.Count, headerRowCount, options, consumerName);
            return new ObjectTableProjection(rows, columns);
        }

        private static void EnsureTableCellLimit(
            int rowCount,
            int columnCount,
            int headerRowCount,
            ObjectFlattenerOptions options,
            string consumerName) {
            long projectedCells = ((long)rowCount + headerRowCount) * columnCount;
            if (projectedCells > options.MaxCells) {
                throw new InvalidDataException(
                    $"{consumerName} exceeds the {options.MaxCells}-cell materialization limit.");
            }
        }
    }

    internal sealed class ObjectTableProjection {
        internal ObjectTableProjection(
            IReadOnlyList<Dictionary<string, object?>> rows,
            IReadOnlyList<string> columns) {
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
        }

        internal IReadOnlyList<Dictionary<string, object?>> Rows { get; }

        internal IReadOnlyList<string> Columns { get; }
    }
}
