using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Data {
    /// <summary>
    /// Provides the shared runtime-row projection used by Office table renderers.
    /// </summary>
    public partial class ObjectFlattener {
        internal ObjectTableProjection FlattenRows<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
            IEnumerable<T> source,
            ObjectFlattenerOptions options,
            string consumerName) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

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
                        discoveredColumns.Add(path);
                    }
                }
            }

            IEnumerable<string> columnCandidates = options.Columns != null && options.Columns.Length > 0
                ? options.Columns
                : discoveredColumns;
            List<string> columns = ResolvePaths(columnCandidates, options);
            return new ObjectTableProjection(rows, columns);
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
