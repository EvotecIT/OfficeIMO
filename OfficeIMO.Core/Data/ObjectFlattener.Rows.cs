using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
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
            bool enforceEmptyProjectionLimits = true,
            int? renderedColumnCountForCellLimit = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            options = options.CreateProjectionSnapshot();
            if (headerRowCount < 0) throw new ArgumentOutOfRangeException(nameof(headerRowCount));
            if (renderedColumnCountForCellLimit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(renderedColumnCountForCellLimit));
            }
            ValidateLimits(options);
            if (options.MaxRows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxRows),
                    "MaxRows must be greater than zero.");
            }

            int knownRowCount = source is IReadOnlyCollection<T> readOnlyCollection
                ? readOnlyCollection.Count
                : source is ICollection<T> collection
                    ? collection.Count
                    : 0;
            if (knownRowCount > options.MaxRows) {
                throw CreateRowLimitException(knownRowCount, options.MaxRows, consumerName);
            }

            var rows = knownRowCount > 0
                ? new List<Dictionary<string, object?>>(Math.Min(knownRowCount, 4096))
                : new List<Dictionary<string, object?>>();
            var discoveredColumns = new List<string>();
            var discoveredColumnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string>? explicitColumns = null;
            bool hasExplicitColumns = options.Columns != null && options.Columns.Length > 0;
            if (hasExplicitColumns) {
                explicitColumns = ResolveExplicitColumns(options, consumerName);
                if (knownRowCount > 0) {
                    EnsureTableCellLimit(
                        knownRowCount,
                        GetRenderedColumnCount(explicitColumns.Count, renderedColumnCountForCellLimit),
                        headerRowCount,
                        options,
                        consumerName);
                    if (renderedColumnCountForCellLimit.HasValue) {
                        EnsureIntermediateCellLimit(
                            knownRowCount,
                            explicitColumns.Count,
                            headerRowCount,
                            options,
                            consumerName);
                    }
                }
            }

            void AddProjectedRow(T item) {
                if (rows.Count >= options.MaxRows) {
                    throw CreateRowLimitException(checked(rows.Count + 1), options.MaxRows, consumerName);
                }
                Dictionary<string, object?> row;
                try {
                    row = FlattenPrepared(item, options);
                } catch (InvalidDataException exception) {
                    if (!TryGetRawColumnLimit(exception, out int requiredColumns, out int columnLimit)) {
                        throw;
                    }
                    throw CreateColumnLimitException(
                        requiredColumns,
                        columnLimit,
                        consumerName,
                        exception);
                }
                rows.Add(row);
                foreach (string path in row.Keys) {
                    if (!string.IsNullOrWhiteSpace(path) && discoveredColumnSet.Add(path)) {
                        if (discoveredColumns.Count >= options.MaxColumns) {
                            throw CreateColumnLimitException(checked(discoveredColumns.Count + 1), options.MaxColumns, consumerName);
                        }
                        discoveredColumns.Add(path);
                    }
                }
                int materializedColumnCount = explicitColumns?.Count ?? discoveredColumns.Count;
                EnsureTableCellLimit(
                    rows.Count,
                    GetRenderedColumnCount(materializedColumnCount, renderedColumnCountForCellLimit),
                    headerRowCount,
                    options,
                    consumerName);
                if (renderedColumnCountForCellLimit.HasValue) {
                    EnsureIntermediateCellLimit(
                        rows.Count,
                        materializedColumnCount,
                        headerRowCount,
                        options,
                        consumerName);
                }
            }

            if (source is IReadOnlyList<T> readOnlyList) {
                for (int index = 0; index < readOnlyList.Count; index++) {
                    AddProjectedRow(readOnlyList[index]);
                }
            } else {
                foreach (T item in source) AddProjectedRow(item);
            }

            if (rows.Count == 0) {
                if (!enforceEmptyProjectionLimits) {
                    return new ObjectTableProjection(rows, new List<string>());
                }
                if (hasExplicitColumns) {
                    explicitColumns = ResolveExplicitColumns(options, consumerName);
                } else if (!ObjectDictionaryAdapter.IsDictionaryType(typeof(T))) {
                    discoveredColumns.AddRange(GetPathsPrepared(typeof(T), options));
                }
            }

            List<string> columns = explicitColumns ?? ResolvePathsPrepared(discoveredColumns, options);
            if (columns.Count > options.MaxColumns) {
                throw CreateColumnLimitException(columns.Count, options.MaxColumns, consumerName);
            }
            EnsureTableCellLimit(
                rows.Count,
                GetRenderedColumnCount(columns.Count, renderedColumnCountForCellLimit),
                headerRowCount,
                options,
                consumerName);
            if (renderedColumnCountForCellLimit.HasValue) {
                EnsureIntermediateCellLimit(rows.Count, columns.Count, headerRowCount, options, consumerName);
            }
            return new ObjectTableProjection(rows, columns);
        }

        private static int GetRenderedColumnCount(int columnCount, int? renderedColumnCountForCellLimit) =>
            renderedColumnCountForCellLimit.HasValue
                ? Math.Min(columnCount, renderedColumnCountForCellLimit.Value)
                : columnCount;

        private static void EnsureIntermediateCellLimit(
            int rowCount,
            int columnCount,
            int headerRowCount,
            ObjectFlattenerOptions options,
            string consumerName) {
            long maximumIntermediateCells = Math.Max(options.MaxCells, ObjectFlattenerOptions.DefaultMaxCells);
            long intermediateCells = ((long)rowCount + headerRowCount) * columnCount;
            if (intermediateCells > maximumIntermediateCells) {
                throw CreateCellLimitException(
                    rowCount,
                    columnCount,
                    headerRowCount,
                    intermediateCells,
                    maximumIntermediateCells,
                    consumerName,
                    "intermediate materialization");
            }
        }

        private List<string> ResolveExplicitColumns(
            ObjectFlattenerOptions options,
            string consumerName) {
            var explicitColumnCandidates = new List<string>(Math.Min(options.Columns!.Length, options.MaxColumns));
            AddExplicitColumnsBounded(explicitColumnCandidates, options, consumerName);
            return ResolvePathsPrepared(explicitColumnCandidates, options);
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
                if (!IsPathSelectedForMaterialization(column, options)) {
                    continue;
                }
                if (columns.Count >= options.MaxColumns) {
                    throw CreateColumnLimitException(checked(columns.Count + 1), options.MaxColumns, consumerName);
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
                throw CreateCellLimitException(
                    rowCount,
                    columnCount,
                    headerRowCount,
                    projectedCells,
                    options.MaxCells,
                    consumerName,
                    "materialization");
            }
        }

        private static InvalidDataException CreateCellLimitException(
            int rowCount,
            int columnCount,
            int headerRowCount,
            long projectedCells,
            long limit,
            string consumerName,
            string limitKind) {
            string rows = headerRowCount == 0
                ? FormatCount(rowCount, "row", "rows")
                : FormatCount(rowCount, "data row", "data rows") + " + "
                    + FormatCount(headerRowCount, "header row", "header rows");
            string requiredCells = projectedCells.ToString(CultureInfo.InvariantCulture);
            string overrideHint = string.Equals(consumerName, "TableFrom", StringComparison.Ordinal)
                ? "For intentionally materialized object rows, raise the limit with configure: options => options.MaxCells = " + requiredCells + ". "
                    + "For fixed-schema data, use the TableFrom(DataTable) overload, which avoids generic object flattening."
                : "If this materialization is intentional, set ObjectFlattenerOptions.MaxCells to at least " + requiredCells + ".";

            return new InvalidDataException(
                consumerName + " requires at least " + requiredCells
                + " cells (" + rows + " x " + columnCount.ToString(CultureInfo.InvariantCulture)
                + " columns), exceeding the " + limit.ToString(CultureInfo.InvariantCulture)
                + "-cell " + limitKind + " limit (MaxCells). " + overrideHint);
        }

        private static InvalidDataException CreateRowLimitException(int requiredRows, int limit, string consumerName) {
            string required = requiredRows.ToString(CultureInfo.InvariantCulture);
            string overrideHint = string.Equals(consumerName, "TableFrom", StringComparison.Ordinal)
                ? "If this materialization is intentional, raise the limit with configure: options => options.MaxRows = " + required + "."
                : "If this materialization is intentional, set ObjectFlattenerOptions.MaxRows to at least " + required + ".";
            return new InvalidDataException(
                consumerName + " requires at least " + FormatCount(requiredRows, "data row", "data rows")
                + ", exceeding the " + limit.ToString(CultureInfo.InvariantCulture)
                + "-row materialization limit (MaxRows). " + overrideHint);
        }

        private static InvalidDataException CreateColumnLimitException(
            int requiredColumns,
            int limit,
            string consumerName,
            Exception? innerException = null) {
            string required = requiredColumns.ToString(CultureInfo.InvariantCulture);
            string overrideHint = string.Equals(consumerName, "TableFrom", StringComparison.Ordinal)
                ? "If this materialization is intentional, raise the limit with configure: options => options.MaxColumns = " + required + "."
                : "If this materialization is intentional, set ObjectFlattenerOptions.MaxColumns to at least " + required + ".";
            string message = consumerName + " requires at least " + FormatCount(requiredColumns, "column", "columns")
                + ", exceeding the " + limit.ToString(CultureInfo.InvariantCulture)
                + "-column materialization limit (MaxColumns). " + overrideHint;
            return innerException == null
                ? new InvalidDataException(message)
                : new InvalidDataException(message, innerException);
        }

        private static string FormatCount(int count, string singular, string plural) {
            return count.ToString(CultureInfo.InvariantCulture) + " " + (count == 1 ? singular : plural);
        }

        internal static InvalidDataException CreateRawColumnLimitException(
            string operation,
            int requiredColumns,
            int limit) {
            var exception = new InvalidDataException(
                operation + " requires at least " + requiredColumns.ToString(CultureInfo.InvariantCulture)
                + " columns, exceeding the " + limit.ToString(CultureInfo.InvariantCulture)
                + "-column limit (MaxColumns). Set ObjectFlattenerOptions.MaxColumns to at least "
                + requiredColumns.ToString(CultureInfo.InvariantCulture) + ".");
            exception.Data["OfficeIMO.RequiredColumns"] = requiredColumns;
            exception.Data["OfficeIMO.MaxColumns"] = limit;
            return exception;
        }

        private static bool TryGetRawColumnLimit(
            InvalidDataException exception,
            out int requiredColumns,
            out int limit) {
            if (exception.Data["OfficeIMO.RequiredColumns"] is int required
                && exception.Data["OfficeIMO.MaxColumns"] is int configuredLimit) {
                requiredColumns = required;
                limit = configuredLimit;
                return true;
            }

            requiredColumns = 0;
            limit = 0;
            return false;
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
