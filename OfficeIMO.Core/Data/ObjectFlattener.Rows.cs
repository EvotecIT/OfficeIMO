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
            int headerRowCount,
            bool enforceEmptyProjectionLimits = true) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (headerRowCount < 0) throw new ArgumentOutOfRangeException(nameof(headerRowCount));
            ValidateLimits(options);

            List<T> items = MaterializeRowsBounded(source, options, consumerName);
            if (items.Count == 0 && !enforceEmptyProjectionLimits) {
                return new ObjectTableProjection(
                    new List<Dictionary<string, object?>>(),
                    new List<string>());
            }

            var rows = new List<Dictionary<string, object?>>(items.Count);
            var discoveredColumns = new List<string>();
            var discoveredColumnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string>? explicitColumns = null;
            if (options.Columns != null && options.Columns.Length > 0) {
                var explicitColumnCandidates = new List<string>(Math.Min(options.Columns.Length, options.MaxColumns));
                AddExplicitColumnsBounded(explicitColumnCandidates, options, consumerName);
                explicitColumns = ResolvePaths(explicitColumnCandidates, options);
            }

            if (items.Count == 0) {
                if (explicitColumns == null && !ObjectDictionaryAdapter.IsDictionaryType(typeof(T))) {
                    discoveredColumns.AddRange(GetPaths(typeof(T), options));
                }
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
                EnsureTableCellLimit(
                    rows.Count,
                    explicitColumns?.Count ?? discoveredColumns.Count,
                    headerRowCount,
                    options,
                    consumerName);
            }

            List<string> columns = explicitColumns ?? ResolvePaths(discoveredColumns, options);
            if (columns.Count > options.MaxColumns) {
                throw new InvalidDataException(
                    $"{consumerName} exceeds the {options.MaxColumns}-column materialization limit.");
            }
            EnsureTableCellLimit(rows.Count, columns.Count, headerRowCount, options, consumerName);
            return new ObjectTableProjection(rows, columns);
        }

        private static void AddExplicitColumnsBounded(
            List<string> columns,
            ObjectFlattenerOptions options,
            string consumerName) {
            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string column in options.Columns!) {
                if (string.IsNullOrWhiteSpace(column) || !added.Add(column)) {
                    continue;
                }
                if (columns.Count >= options.MaxColumns) {
                    throw new InvalidDataException(
                        $"{consumerName} exceeds the {options.MaxColumns}-column materialization limit.");
                }
                columns.Add(column);
            }
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
