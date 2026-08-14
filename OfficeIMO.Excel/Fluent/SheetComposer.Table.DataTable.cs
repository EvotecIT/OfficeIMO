using System.Data;
using System.IO;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.Fluent {
    public sealed partial class SheetComposer {
        /// <summary>
        /// Default maximum number of cells written by the fixed-schema <see cref="DataTable"/> overload,
        /// including the header row.
        /// </summary>
        public const long DefaultDataTableMaxCells = 2_000_000L;

        /// <summary>
        /// Renders a fixed-schema <see cref="DataTable"/> as an editable Excel table without generic object flattening.
        /// The DataTable schema order is preserved unless the projection configuration selects another order.
        /// </summary>
        /// <param name="table">Source table whose rows and schema are already materialized.</param>
        /// <param name="title">Optional section title and table name.</param>
        /// <param name="configure">
        /// Optional column selection, exclusion, ordering, header, and materialization-limit configuration.
        /// The fixed-schema overload defaults <see cref="ObjectFlattenerOptions.MaxCells"/> to
        /// <see cref="DefaultDataTableMaxCells"/> instead of the lower generic object-flattening limit.
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

            if (table.Rows.Count == 0) {
                if (!string.IsNullOrWhiteSpace(title)) Section(title!);
                Sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            var options = new ObjectFlattenerOptions {
                MaxCells = DefaultDataTableMaxCells
            };
            configure?.Invoke(options);
            if (options.MaxRows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxRows), "MaxRows must be greater than zero.");
            }
            if (options.MaxCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxCells), "MaxCells must be greater than zero.");
            }
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
                paths = ResolveExplicitDataTablePaths(flattener, options, sourceColumnNames);
            } else {
                paths = ResolveProjectedDataTablePaths(
                    options,
                    sourceColumnNames,
                    configuredMaxColumns);
            }
            if (paths.Count == 0) {
                if (!string.IsNullOrWhiteSpace(title)) Section(title!);
                Sheet.Cell(_row, 1, "(no tabular columns for the DataTable schema)");
                _row++;
                return $"A{_row - 1}:A{_row - 1}";
            }

            int headerRow = checked(_row + (string.IsNullOrWhiteSpace(title) ? 0 : 1));
            int lastRow = checked(headerRow + table.Rows.Count);
            if (lastRow > A1.MaxRows) {
                throw new InvalidDataException(
                    $"TableFrom(DataTable) requires rows {headerRow} through {lastRow}, exceeding Excel's {A1.MaxRows}-row worksheet limit. "
                    + "Split the data across multiple worksheets or start the table earlier; Excel's worksheet row limit cannot be overridden.");
            }

            EnsureDataTableMaterializationLimits(table.Rows.Count, paths.Count, options);

            if (!string.IsNullOrWhiteSpace(title)) Section(title!);
            List<string> headers = BuildTransformedHeaders(paths, options);
            EnsureUniqueTableHeaders(headers);
            var source = new DataTableTabularRowSource(table, paths, headers);
            string range = Sheet.InsertTabularRowSourceAsTable(
                source,
                startRow: headerRow,
                startColumn: 1,
                includeHeaders: true,
                tableName: title ?? "Table",
                style: style,
                includeAutoFilter: autoFilter);
            return CompleteTable(
                range,
                paths,
                headers,
                headerRow,
                lastRow,
                style,
                freezeHeaderRow,
                visuals,
                pathsRepresentFlattenedObjects: false);
        }

        private static List<string> ResolveExplicitDataTablePaths(
            ObjectFlattener flattener,
            ObjectFlattenerOptions options,
            IReadOnlyList<string> sourceColumnNames) {
            var paths = new List<string>();
            var added = new HashSet<string>(StringComparer.Ordinal);
            var selectedSourceColumns = ResolveDataTableFilterSelection(sourceColumnNames, options);
            var selectedSourceColumnSet = new HashSet<string>(selectedSourceColumns, StringComparer.Ordinal);
            foreach (string candidate in options.Columns!) {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                string canonicalCandidate = ResolveExplicitDataTableColumn(
                    sourceColumnNames,
                    candidate,
                    nameof(options.Columns)) ?? candidate;
                if (!added.Add(canonicalCandidate)) continue;
                bool isSourceColumn = sourceColumnNames.Contains(canonicalCandidate, StringComparer.Ordinal);
                if (isSourceColumn) {
                    if (!selectedSourceColumnSet.Contains(canonicalCandidate)) continue;
                } else {
                    List<string> selected = flattener.ResolvePaths(new[] { canonicalCandidate }, options);
                    if (selected.Count == 0) continue;
                }
                if (paths.Count >= options.MaxColumns) {
                    EnsureDataTableColumnLimit(checked(paths.Count + 1), options.MaxColumns);
                }
                paths.Add(canonicalCandidate);
            }
            return paths;
        }

        private static string? ResolveExplicitDataTableColumn(
            IReadOnlyList<string> paths,
            string candidate,
            string optionName) {
            string? exactPath = paths.FirstOrDefault(path =>
                string.Equals(path, candidate, StringComparison.Ordinal));
            if (exactPath != null) return exactPath;

            List<string> insensitivePaths = paths
                .Where(path => string.Equals(path, candidate, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (insensitivePaths.Count == 1) return insensitivePaths[0];
            if (insensitivePaths.Count > 1) {
                throw CreateAmbiguousDataTableRuleException(candidate, optionName);
            }
            return null;
        }

        private static List<string> ResolveProjectedDataTablePaths(
            ObjectFlattenerOptions options,
            IReadOnlyList<string> sourceColumnNames,
            int configuredMaxColumns) {
            List<string> selected = ResolveDataTableFilterSelection(sourceColumnNames, options);
            EnsureDataTableColumnLimit(selected.Count, configuredMaxColumns);
            return ApplyDataTableOrdering(selected, options);
        }

        private static List<string> ResolveDataTableFilterSelection(
            IReadOnlyList<string> sourceColumnNames,
            ObjectFlattenerOptions options) {
            var selected = new List<string>(sourceColumnNames);
            var ignored = ResolveDataTableFilterMatches(
                sourceColumnNames,
                options.Ignore,
                nameof(options.Ignore),
                prefixMatch: true,
                removeAllCaseVariants: true);
            var excluded = ResolveDataTableFilterMatches(
                sourceColumnNames,
                options.ExcludeProperties,
                nameof(options.ExcludeProperties),
                prefixMatch: false,
                removeAllCaseVariants: true);
            var removed = new HashSet<string>(ignored, StringComparer.Ordinal);
            removed.UnionWith(excluded);
            selected.RemoveAll(removed.Contains);

            if (options.IncludeProperties.Length > 0) {
                var included = new HashSet<string>(
                    ResolveDataTableFilterMatches(
                        sourceColumnNames,
                        options.IncludeProperties,
                        nameof(options.IncludeProperties),
                        prefixMatch: false,
                        removeAllCaseVariants: false),
                    StringComparer.Ordinal);
                selected.RemoveAll(path => !included.Contains(path));
            }

            return selected;
        }

        private static List<string> ResolveDataTableFilterMatches(
            IReadOnlyList<string> paths,
            IEnumerable<string>? rules,
            string optionName,
            bool prefixMatch,
            bool removeAllCaseVariants) {
            var result = new List<string>();
            var added = new HashSet<string>(StringComparer.Ordinal);
            foreach (string rule in rules ?? Array.Empty<string>()) {
                if (string.IsNullOrWhiteSpace(rule)) continue;

                if (removeAllCaseVariants) {
                    foreach (string path in paths.Where(path =>
                        (prefixMatch
                            ? path.StartsWith(rule, StringComparison.OrdinalIgnoreCase)
                            : string.Equals(path, rule, StringComparison.OrdinalIgnoreCase))
                        || (!prefixMatch
                            && string.Equals(GetLastSegment(path), rule, StringComparison.OrdinalIgnoreCase)))) {
                        if (added.Add(path)) result.Add(path);
                    }
                    continue;
                }

                List<string> exactPaths = paths
                    .Where(path => prefixMatch
                        ? path.StartsWith(rule, StringComparison.Ordinal)
                        : string.Equals(path, rule, StringComparison.Ordinal))
                    .ToList();
                if (exactPaths.Count > 0) {
                    foreach (string path in exactPaths) if (added.Add(path)) result.Add(path);
                    continue;
                }

                if (!prefixMatch) {
                    List<string> exactSegments = paths
                        .Where(path => string.Equals(GetLastSegment(path), rule, StringComparison.Ordinal))
                        .ToList();
                    if (exactSegments.Count > 0) {
                        foreach (string path in exactSegments) if (added.Add(path)) result.Add(path);
                        continue;
                    }
                }

                List<string> insensitivePaths = paths
                    .Where(path => prefixMatch
                        ? path.StartsWith(rule, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(path, rule, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (insensitivePaths.Count > 0) {
                    EnsureUnambiguousDataTableFilterFallback(insensitivePaths, rule, optionName, prefixMatch);
                    foreach (string path in insensitivePaths) if (added.Add(path)) result.Add(path);
                    continue;
                }

                if (!prefixMatch) {
                    List<string> insensitiveSegments = paths
                        .Where(path => string.Equals(GetLastSegment(path), rule, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    if (insensitiveSegments.Count > 0) {
                        EnsureUnambiguousDataTableFilterFallback(insensitiveSegments, rule, optionName, prefixMatch: false, compareLastSegment: true);
                        foreach (string path in insensitiveSegments) if (added.Add(path)) result.Add(path);
                    }
                }
            }
            return result;
        }

        private static void EnsureUnambiguousDataTableFilterFallback(
            IReadOnlyList<string> matches,
            string rule,
            string optionName,
            bool prefixMatch,
            bool compareLastSegment = false) {
            var spellings = new HashSet<string>(StringComparer.Ordinal);
            foreach (string match in matches) {
                string spelling = compareLastSegment
                    ? GetLastSegment(match)
                    : prefixMatch ? match.Substring(0, Math.Min(rule.Length, match.Length)) : match;
                spellings.Add(spelling);
            }
            if (spellings.Count > 1) throw CreateAmbiguousDataTableRuleException(rule, optionName);
        }

        private static List<string> ApplyDataTableOrdering(
            IReadOnlyList<string> paths,
            ObjectFlattenerOptions options) {
            if (paths.Count == 0) return new List<string>();

            var pinnedFirst = ResolveDataTableRules(paths, options.PinnedFirst, nameof(options.PinnedFirst));
            var pinnedLast = ResolveDataTableRules(paths, options.PinnedLast, nameof(options.PinnedLast));
            var firstSet = new HashSet<string>(pinnedFirst, StringComparer.Ordinal);
            var lastSet = new HashSet<string>(pinnedLast, StringComparer.Ordinal);
            lastSet.ExceptWith(firstSet);

            var priorities = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, int> priority in options.PropertyPriority) {
                string? path = ResolveDataTableRule(paths, priority.Key, nameof(options.PropertyPriority));
                if (path != null && !priorities.ContainsKey(path)) priorities[path] = priority.Value;
            }

            var remaining = paths
                .Select((path, index) => new {
                    Path = path,
                    Index = index,
                    Priority = priorities.TryGetValue(path, out int priority) ? priority : 0
                })
                .Where(item => !firstSet.Contains(item.Path) && !lastSet.Contains(item.Path))
                .OrderBy(item => item.Priority)
                .ThenBy(item => item.Index)
                .Select(item => item.Path);

            var result = new List<string>(paths.Count);
            result.AddRange(pinnedFirst);
            result.AddRange(remaining);
            result.AddRange(pinnedLast.Where(path => !firstSet.Contains(path)));
            return result;
        }

        private static List<string> ResolveDataTableRules(
            IReadOnlyList<string> paths,
            IEnumerable<string>? rules,
            string optionName) {
            var result = new List<string>();
            var added = new HashSet<string>(StringComparer.Ordinal);
            foreach (string rule in rules ?? Array.Empty<string>()) {
                string? path = ResolveDataTableRule(paths, rule, optionName);
                if (path != null && added.Add(path)) result.Add(path);
            }
            return result;
        }

        private static string? ResolveDataTableRule(
            IReadOnlyList<string> paths,
            string rule,
            string optionName) {
            if (string.IsNullOrWhiteSpace(rule)) return null;

            string? exactPath = paths.FirstOrDefault(path => string.Equals(path, rule, StringComparison.Ordinal));
            if (exactPath != null) return exactPath;

            List<string> exactSegments = paths
                .Where(path => string.Equals(GetLastSegment(path), rule, StringComparison.Ordinal))
                .ToList();
            if (exactSegments.Count == 1) return exactSegments[0];
            if (exactSegments.Count > 1) throw CreateAmbiguousDataTableRuleException(rule, optionName);

            List<string> insensitivePaths = paths
                .Where(path => string.Equals(path, rule, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (insensitivePaths.Count == 1) return insensitivePaths[0];
            if (insensitivePaths.Count > 1) throw CreateAmbiguousDataTableRuleException(rule, optionName);

            List<string> insensitiveSegments = paths
                .Where(path => string.Equals(GetLastSegment(path), rule, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (insensitiveSegments.Count == 1) return insensitiveSegments[0];
            if (insensitiveSegments.Count > 1) throw CreateAmbiguousDataTableRuleException(rule, optionName);
            return null;
        }

        private static string GetLastSegment(string path) {
            int separator = path.LastIndexOf('.');
            return separator >= 0 ? path.Substring(separator + 1) : path;
        }

        private static InvalidDataException CreateAmbiguousDataTableRuleException(string rule, string optionName) =>
            new InvalidDataException(
                $"DataTable column rule '{rule}' in {optionName} is ambiguous because the schema contains case-distinct or repeated matches. "
                + "Use the exact column casing by specifying the exact full column name and casing in the projection rule.");

        private static void EnsureDataTableColumnLimit(int requiredColumns, int configuredMaxColumns) {
            if (requiredColumns <= configuredMaxColumns && requiredColumns <= A1.MaxColumns) return;
            if (requiredColumns > A1.MaxColumns) {
                throw CreateExcelColumnLimitException(requiredColumns);
            }

            throw new InvalidDataException(
                $"TableFrom(DataTable) requires at least {requiredColumns} columns, exceeding the {configuredMaxColumns}-column materialization limit (MaxColumns). "
                + $"If this projection is intentional, raise the limit with configure: options => options.MaxColumns = {requiredColumns}.");
        }

        private static void EnsureDataTableMaterializationLimits(
            int requiredRows,
            int requiredColumns,
            ObjectFlattenerOptions options) {
            if (requiredRows > options.MaxRows) {
                throw new InvalidDataException(
                    $"TableFrom(DataTable) requires {requiredRows} data rows, exceeding the {options.MaxRows}-row materialization limit (MaxRows). "
                    + $"If this materialization is intentional, raise the limit with configure: options => options.MaxRows = {requiredRows}.");
            }

            long requiredCells = checked(((long)requiredRows + 1L) * requiredColumns);
            if (requiredCells > options.MaxCells) {
                throw new InvalidDataException(
                    $"TableFrom(DataTable) requires {requiredCells} cells including the header row, exceeding the {options.MaxCells}-cell "
                    + "fixed-schema materialization limit (MaxCells). "
                    + $"If this materialization is intentional, raise the limit with configure: options => options.MaxCells = {requiredCells}.");
            }
        }

        private static InvalidDataException CreateExcelColumnLimitException(int requiredColumns) =>
            new InvalidDataException(
                $"TableFrom(DataTable) requires at least {requiredColumns} columns, exceeding Excel's {A1.MaxColumns}-column worksheet limit. "
                + "Select fewer columns or split the data across multiple worksheets; Excel's worksheet column limit cannot be overridden.");
    }
}
