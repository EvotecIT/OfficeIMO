using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Workbook or worksheet formula inspection result.
    /// </summary>
    public sealed class ExcelFormulaInspection {
        internal ExcelFormulaInspection(IReadOnlyList<ExcelFormulaCellInfo> formulas) {
            Formulas = formulas;
            TotalFormulas = formulas.Count;
            SupportedFormulas = formulas.Count(formula => formula.IsSupportedByOfficeIMO);
            UnsupportedFormulas = TotalFormulas - SupportedFormulas;
            MissingCachedResults = formulas.Count(formula => !formula.HasCachedValue);
            DirtyFormulas = formulas.Count(formula => formula.IsDirty);
            DependencyGraph = ExcelFormulaDependencyGraph.Create(formulas);
            DependencyIssueCount = formulas.Sum(formula => formula.DependencyIssues.Count) +
                DependencyGraph.CircularReferenceCount +
                (DependencyGraph.AnalysisTruncated ? 1 : 0);
        }

        /// <summary>Formula cells discovered in workbook order.</summary>
        public IReadOnlyList<ExcelFormulaCellInfo> Formulas { get; }

        /// <summary>Total formula count.</summary>
        public int TotalFormulas { get; }

        /// <summary>Formula count supported by OfficeIMO's lightweight evaluator.</summary>
        public int SupportedFormulas { get; }

        /// <summary>Formula count that must be preserved, cached, or recalculated by Excel.</summary>
        public int UnsupportedFormulas { get; }

        /// <summary>Formula cells without cached results.</summary>
        public int MissingCachedResults { get; }

        /// <summary>Formula cells marked dirty for recalculation.</summary>
        public int DirtyFormulas { get; }

        /// <summary>Total dependency issues found across formula cells.</summary>
        public int DependencyIssueCount { get; }

        /// <summary>Workbook-level graph of formula cells, direct dependencies, and formula dependents.</summary>
        public ExcelFormulaDependencyGraph DependencyGraph { get; }

        /// <summary>True when every formula can be evaluated by OfficeIMO's lightweight evaluator.</summary>
        public bool AllSupported => TotalFormulas == SupportedFormulas;

        /// <summary>True when every formula has a cached result.</summary>
        public bool AllHaveCachedResults => MissingCachedResults == 0;

        /// <summary>True when any formula dependency issue was detected.</summary>
        public bool HasDependencyIssues => DependencyIssueCount > 0;

        /// <summary>Describes the formula patterns supported by OfficeIMO's lightweight evaluator.</summary>
        public ExcelFormulaCapabilities Capabilities => ExcelFormulaCapabilities.Current;

        /// <summary>
        /// Throws when any formula is outside OfficeIMO's lightweight evaluator support.
        /// </summary>
        public ExcelFormulaInspection EnsureAllSupported() {
            if (UnsupportedFormulas > 0) {
                var unsupported = Formulas
                    .Where(formula => !formula.IsSupportedByOfficeIMO)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray();
                throw new InvalidOperationException("Unsupported formulas: " + string.Join(", ", unsupported));
            }

            return this;
        }

        /// <summary>
        /// Throws when any formula lacks a cached result.
        /// </summary>
        public ExcelFormulaInspection EnsureAllHaveCachedResults() {
            if (MissingCachedResults > 0) {
                var missing = Formulas
                    .Where(formula => !formula.HasCachedValue)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray();
                throw new InvalidOperationException("Formula cells without cached results: " + string.Join(", ", missing));
            }

            return this;
        }

        /// <summary>
        /// Throws when any formula dependency issue is detected.
        /// </summary>
        public ExcelFormulaInspection EnsureNoDependencyIssues() {
            if (DependencyIssueCount > 0) {
                var issues = Formulas
                    .Where(formula => formula.DependencyIssues.Count > 0)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}: {string.Join("; ", formula.DependencyIssues)}")
                    .Concat(DependencyGraph.CircularReferences.Select(circular => $"Circular reference: {string.Join(" -> ", circular.References)}"))
                    .Concat(DependencyGraph.AnalysisTruncated
                        ? new[] { DependencyGraph.AnalysisTruncationReason! }
                        : Array.Empty<string>())
                    .ToArray();
                throw new InvalidOperationException("Formula dependency issues: " + string.Join(", ", issues));
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of formula support and cache status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Excel Formula Inspection");
            builder.AppendLine();
            builder.AppendLine($"Total formulas: {TotalFormulas}");
            builder.AppendLine($"Supported formulas: {SupportedFormulas}");
            builder.AppendLine($"Unsupported formulas: {UnsupportedFormulas}");
            builder.AppendLine($"Missing cached results: {MissingCachedResults}");
            builder.AppendLine($"Dirty formulas: {DirtyFormulas}");
            builder.AppendLine($"Dependency issues: {DependencyIssueCount}");
            builder.AppendLine($"Maximum formula dependency depth: {DependencyGraph.MaximumDependencyDepth}");
            builder.AppendLine($"Circular reference groups: {DependencyGraph.CircularReferenceCount}");
            builder.AppendLine($"Dependency graph analysis complete: {(DependencyGraph.AnalysisTruncated ? "no" : "yes")}");
            builder.AppendLine();
            builder.AppendLine("| Sheet | Cell | Formula | Supported | Cached | Dirty | Dependencies | Dependency issues | Reason |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- |");

            foreach (ExcelFormulaCellInfo formula in Formulas) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(formula.SheetName));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.CellReference));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.Formula));
                builder.Append(" | ");
                builder.Append(formula.IsSupportedByOfficeIMO ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(formula.HasCachedValue ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(formula.IsDirty ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", formula.Dependencies)));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", formula.DependencyIssues)));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.UnsupportedReason ?? string.Empty));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }

    /// <summary>
    /// Workbook-level dependency graph built from formula inspection metadata.
    /// </summary>
    public sealed class ExcelFormulaDependencyGraph {
        private const int MaximumFormulaDependenciesPerCell = 4096;
        private const int MaximumFormulaGraphEdges = 100_000;
        private readonly Dictionary<string, ExcelFormulaDependencyNode> _nodesByReference;

        private ExcelFormulaDependencyGraph(
            IReadOnlyList<ExcelFormulaDependencyNode> nodes,
            IReadOnlyList<ExcelFormulaCircularReference> circularReferences,
            int maximumDependencyDepth,
            string? analysisTruncationReason) {
            Nodes = nodes;
            CircularReferences = circularReferences;
            MaximumDependencyDepth = maximumDependencyDepth;
            AnalysisTruncationReason = analysisTruncationReason;
            _nodesByReference = nodes.ToDictionary(node => node.Reference, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>Formula nodes in workbook order.</summary>
        public IReadOnlyList<ExcelFormulaDependencyNode> Nodes { get; }

        /// <summary>Total formula nodes in the graph.</summary>
        public int NodeCount => Nodes.Count;

        /// <summary>Total direct formula-to-formula dependency edges.</summary>
        public int EdgeCount => Nodes.Sum(node => node.FormulaDependencies.Count);

        /// <summary>Longest resolvable formula-to-formula dependency chain, counting formula cells.</summary>
        public int MaximumDependencyDepth { get; }

        /// <summary>Circular formula-reference groups detected in the workbook.</summary>
        public IReadOnlyList<ExcelFormulaCircularReference> CircularReferences { get; }

        /// <summary>Total circular formula-reference groups.</summary>
        public int CircularReferenceCount => CircularReferences.Count;

        /// <summary>True when one or more circular formula-reference groups were detected.</summary>
        public bool HasCircularReferences => CircularReferenceCount > 0;

        /// <summary>True when resource limits stopped dependency-edge discovery before the graph was complete.</summary>
        public bool AnalysisTruncated => AnalysisTruncationReason != null;

        /// <summary>Reason dependency-edge discovery was truncated, or null when the graph is complete.</summary>
        public string? AnalysisTruncationReason { get; }

        /// <summary>True when at least one formula node has dependency diagnostics.</summary>
        public bool HasDependencyIssues => AnalysisTruncated || HasCircularReferences || Nodes.Any(node => node.HasDependencyIssues);

        internal static ExcelFormulaDependencyGraph Create(IReadOnlyList<ExcelFormulaCellInfo> formulas) {
            var builders = new Dictionary<string, ExcelFormulaDependencyNodeBuilder>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelFormulaCellInfo formula in formulas) {
                string reference = FormatReference(formula.SheetName, formula.CellReference);
                if (builders.ContainsKey(reference)) {
                    continue;
                }

                builders.Add(reference, new ExcelFormulaDependencyNodeBuilder(
                    reference,
                    formula.SheetName,
                    NormalizeCellReference(formula.CellReference),
                    formula.Dependencies,
                    formula.DependencyIssues));
            }

            var formulaCellsBySheet = new Dictionary<string, List<FormulaCellIndexEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelFormulaDependencyNodeBuilder builder in builders.Values) {
                var cell = A1.ParseCellRef(builder.CellReference);
                if (cell.Row <= 0 || cell.Col <= 0) {
                    continue;
                }

                if (!formulaCellsBySheet.TryGetValue(builder.SheetName, out List<FormulaCellIndexEntry>? entries)) {
                    entries = new List<FormulaCellIndexEntry>();
                    formulaCellsBySheet.Add(builder.SheetName, entries);
                }

                entries.Add(new FormulaCellIndexEntry(cell.Row, cell.Col, builder.Reference));
            }
            foreach (List<FormulaCellIndexEntry> entries in formulaCellsBySheet.Values) {
                entries.Sort((left, right) => {
                    int rowComparison = left.Row.CompareTo(right.Row);
                    return rowComparison != 0 ? rowComparison : left.Column.CompareTo(right.Column);
                });
            }

            int remainingGraphEdges = MaximumFormulaGraphEdges;
            string? analysisTruncationReason = null;
            foreach (ExcelFormulaCellInfo formula in formulas) {
                string sourceReference = FormatReference(formula.SheetName, formula.CellReference);
                int formulaDependencyCount = 0;
                bool formulaTruncated = false;
                bool graphTruncated = false;
                foreach (string dependency in formula.Dependencies) {
                    using IEnumerator<string> targets = FindCoveredFormulaReferences(
                        dependency,
                        builders,
                        formulaCellsBySheet).GetEnumerator();
                    while (targets.MoveNext()) {
                        if (remainingGraphEdges == 0) {
                            graphTruncated = true;
                            analysisTruncationReason =
                                $"Formula dependency graph exceeded the global limit of {MaximumFormulaGraphEdges} edges.";
                            break;
                        }
                        if (formulaDependencyCount == MaximumFormulaDependenciesPerCell) {
                            formulaTruncated = true;
                            analysisTruncationReason ??=
                                $"One or more formulas exceeded the limit of {MaximumFormulaDependenciesPerCell} formula dependencies.";
                            break;
                        }
                        string targetReference = targets.Current;
                        if (builders.TryGetValue(sourceReference, out ExcelFormulaDependencyNodeBuilder? sourceNode)
                            && builders.TryGetValue(targetReference, out ExcelFormulaDependencyNodeBuilder? dependencyNode)) {
                            if (sourceNode.FormulaDependencies.Add(targetReference)) {
                                dependencyNode.Dependents.Add(sourceReference);
                                formulaDependencyCount++;
                                remainingGraphEdges--;
                            }
                        }
                    }
                    if (graphTruncated || formulaTruncated) break;
                }
                if (graphTruncated) break;
            }

            ExcelFormulaDependencyAnalysisResult analysis = ExcelFormulaDependencyAnalysis.Analyze(
                builders.ToDictionary(
                    pair => pair.Key,
                    pair => (IReadOnlyCollection<string>)pair.Value.FormulaDependencies,
                    StringComparer.OrdinalIgnoreCase));
            var nodes = builders.Values
                .Select(builder => builder.ToNode(
                    analysis.Depths.TryGetValue(builder.Reference, out int depth) ? depth : (int?)null,
                    analysis.CircularReferences.Contains(builder.Reference)))
                .ToList();
            var circularReferences = analysis.CircularReferenceGroups
                .Select(references => new ExcelFormulaCircularReference(references))
                .ToList();
            return new ExcelFormulaDependencyGraph(
                nodes,
                circularReferences,
                analysis.MaximumDepth,
                analysisTruncationReason);
        }

        /// <summary>
        /// Finds a formula dependency node by worksheet name and cell reference.
        /// </summary>
        public ExcelFormulaDependencyNode? FindNode(string sheetName, string cellReference) {
            string reference = FormatReference(sheetName, cellReference);
            return _nodesByReference.TryGetValue(reference, out ExcelFormulaDependencyNode? node) ? node : null;
        }

        /// <summary>
        /// Returns a compact Markdown report of formula dependencies and dependents.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Excel Formula Dependency Graph");
            builder.AppendLine();
            builder.AppendLine($"Formula nodes: {NodeCount}");
            builder.AppendLine($"Formula dependency edges: {EdgeCount}");
            builder.AppendLine($"Maximum dependency depth: {MaximumDependencyDepth}");
            builder.AppendLine($"Circular reference groups: {CircularReferenceCount}");
            builder.AppendLine($"Analysis complete: {(AnalysisTruncated ? "no" : "yes")}");
            if (AnalysisTruncationReason != null) builder.AppendLine($"Analysis issue: {AnalysisTruncationReason}");
            builder.AppendLine();
            builder.AppendLine("| Formula | Dependencies | Formula dependencies | Dependents | Depth | Circular | Dependency issues |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (ExcelFormulaDependencyNode node in Nodes) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(node.Reference));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", node.Dependencies)));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", node.FormulaDependencies)));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", node.Dependents)));
                builder.Append(" | ");
                builder.Append(node.DependencyDepth?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "unresolved");
                builder.Append(" | ");
                builder.Append(node.IsCircular ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", node.DependencyIssues)));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static IEnumerable<string> FindCoveredFormulaReferences(
            string dependency,
            IReadOnlyDictionary<string, ExcelFormulaDependencyNodeBuilder> builders,
            IReadOnlyDictionary<string, List<FormulaCellIndexEntry>> formulaCellsBySheet) {
            if (!TrySplitQualifiedReference(dependency, out string dependencySheet, out string dependencyAddress)) {
                yield break;
            }

            string normalizedDependency = dependencyAddress.Replace("$", string.Empty);
            if (A1.TryParseRange(normalizedDependency, out int r1, out int c1, out int r2, out int c2)) {
                if (!formulaCellsBySheet.TryGetValue(dependencySheet, out List<FormulaCellIndexEntry>? entries)) {
                    yield break;
                }

                int index = FindFirstFormulaRow(entries, r1);
                for (; index < entries.Count && entries[index].Row <= r2; index++) {
                    FormulaCellIndexEntry entry = entries[index];
                    if (entry.Column >= c1 && entry.Column <= c2) {
                        yield return entry.Reference;
                    }
                }

                yield break;
            }

            if (A1.TryParseWholeColumnRange(normalizedDependency, out c1, out c2)) {
                if (!formulaCellsBySheet.TryGetValue(dependencySheet, out List<FormulaCellIndexEntry>? entries)) {
                    yield break;
                }

                foreach (FormulaCellIndexEntry entry in entries) {
                    if (entry.Column >= c1 && entry.Column <= c2) {
                        yield return entry.Reference;
                    }
                }

                yield break;
            }

            if (A1.TryParseWholeRowRange(normalizedDependency, out r1, out r2)) {
                if (!formulaCellsBySheet.TryGetValue(dependencySheet, out List<FormulaCellIndexEntry>? entries)) {
                    yield break;
                }

                int index = FindFirstFormulaRow(entries, r1);
                for (; index < entries.Count && entries[index].Row <= r2; index++) {
                    yield return entries[index].Reference;
                }

                yield break;
            }

            var dependencyCell = A1.ParseCellRef(normalizedDependency);
            if (dependencyCell.Row <= 0 || dependencyCell.Col <= 0) {
                yield break;
            }

            string targetReference = FormatReference(
                dependencySheet,
                A1.CellReference(dependencyCell.Row, dependencyCell.Col));
            if (builders.ContainsKey(targetReference)) {
                yield return targetReference;
            }
        }

        private static int FindFirstFormulaRow(IReadOnlyList<FormulaCellIndexEntry> entries, int minimumRow) {
            int low = 0;
            int high = entries.Count;
            while (low < high) {
                int middle = low + ((high - low) / 2);
                if (entries[middle].Row < minimumRow) {
                    low = middle + 1;
                } else {
                    high = middle;
                }
            }

            return low;
        }

        private static bool TrySplitQualifiedReference(string reference, out string sheetName, out string address) {
            sheetName = string.Empty;
            address = string.Empty;
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            int separator = reference.LastIndexOf('!');
            if (separator <= 0 || separator == reference.Length - 1) {
                return false;
            }

            sheetName = reference.Substring(0, separator).Trim();
            address = reference.Substring(separator + 1).Trim();
            return sheetName.Length > 0 && address.Length > 0;
        }

        private static string FormatReference(string sheetName, string cellReference) {
            return sheetName + "!" + NormalizeCellReference(cellReference);
        }

        private static string NormalizeCellReference(string cellReference) {
            string normalized = (cellReference ?? string.Empty).Trim().Replace("$", string.Empty);
            var parsed = A1.ParseCellRef(normalized);
            return parsed.Row > 0 && parsed.Col > 0
                ? A1.CellReference(parsed.Row, parsed.Col)
                : normalized;
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }

        private sealed class FormulaCellIndexEntry {
            internal FormulaCellIndexEntry(int row, int column, string reference) {
                Row = row;
                Column = column;
                Reference = reference;
            }

            internal int Row { get; }
            internal int Column { get; }
            internal string Reference { get; }
        }

        private sealed class ExcelFormulaDependencyNodeBuilder {
            internal ExcelFormulaDependencyNodeBuilder(
                string reference,
                string sheetName,
                string cellReference,
                IReadOnlyList<string> dependencies,
                IReadOnlyList<string> dependencyIssues) {
                Reference = reference;
                SheetName = sheetName;
                CellReference = cellReference;
                Dependencies = dependencies;
                DependencyIssues = dependencyIssues;
                FormulaDependencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Dependents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            internal string Reference { get; }
            internal string SheetName { get; }
            internal string CellReference { get; }
            internal IReadOnlyList<string> Dependencies { get; }
            internal IReadOnlyList<string> DependencyIssues { get; }
            internal HashSet<string> FormulaDependencies { get; }
            internal HashSet<string> Dependents { get; }

            internal ExcelFormulaDependencyNode ToNode(int? dependencyDepth, bool isCircular) {
                return new ExcelFormulaDependencyNode(
                    Reference,
                    SheetName,
                    CellReference,
                    Dependencies,
                    FormulaDependencies.OrderBy(dependency => dependency, StringComparer.OrdinalIgnoreCase).ToList(),
                    Dependents.OrderBy(dependent => dependent, StringComparer.OrdinalIgnoreCase).ToList(),
                    DependencyIssues,
                    dependencyDepth,
                    isCircular);
            }
        }
    }

    /// <summary>
    /// Formula cell node in a workbook-level dependency graph.
    /// </summary>
    public sealed class ExcelFormulaDependencyNode {
        internal ExcelFormulaDependencyNode(
            string reference,
            string sheetName,
            string cellReference,
            IReadOnlyList<string> dependencies,
            IReadOnlyList<string> formulaDependencies,
            IReadOnlyList<string> dependents,
            IReadOnlyList<string> dependencyIssues,
            int? dependencyDepth,
            bool isCircular) {
            Reference = reference;
            SheetName = sheetName;
            CellReference = cellReference;
            Dependencies = dependencies;
            FormulaDependencies = formulaDependencies;
            Dependents = dependents;
            DependencyIssues = dependencyIssues;
            DependencyDepth = dependencyDepth;
            IsCircular = isCircular;
        }

        /// <summary>Qualified formula cell reference, such as Sheet1!A1.</summary>
        public string Reference { get; }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 cell reference.</summary>
        public string CellReference { get; }

        /// <summary>Direct A1/range dependencies detected for this formula.</summary>
        public IReadOnlyList<string> Dependencies { get; }

        /// <summary>Formula cells directly referenced by this formula cell.</summary>
        public IReadOnlyList<string> FormulaDependencies { get; }

        /// <summary>Formula cells that directly depend on this formula cell.</summary>
        public IReadOnlyList<string> Dependents { get; }

        /// <summary>Dependency diagnostics associated with this formula cell.</summary>
        public IReadOnlyList<string> DependencyIssues { get; }

        /// <summary>
        /// Resolvable dependency depth for this formula cell, or null when a circular dependency prevents a finite depth.
        /// </summary>
        public int? DependencyDepth { get; }

        /// <summary>True when this formula cell participates directly in a circular reference.</summary>
        public bool IsCircular { get; }

        /// <summary>True when dependency diagnostics found one or more issues.</summary>
        public bool HasDependencyIssues => IsCircular || DependencyIssues.Count > 0;
    }

    /// <summary>
    /// A strongly connected group of formula cells that forms a circular reference.
    /// </summary>
    public sealed class ExcelFormulaCircularReference {
        internal ExcelFormulaCircularReference(IReadOnlyList<string> references) {
            References = references;
        }

        /// <summary>Qualified formula cell references in ordinal order.</summary>
        public IReadOnlyList<string> References { get; }
    }

    /// <summary>
    /// Formula metadata for a single worksheet cell.
    /// </summary>
    public sealed class ExcelFormulaCellInfo {
        internal ExcelFormulaCellInfo(
            string sheetName,
            string cellReference,
            string formula,
            string? cachedValue,
            bool isDirty,
            bool isSupportedByOfficeIMO,
            string? unsupportedReason,
            IReadOnlyList<string>? dependencies = null,
            IReadOnlyList<string>? dependencyIssues = null) {
            SheetName = sheetName;
            CellReference = cellReference;
            Formula = formula;
            CachedValue = cachedValue;
            IsDirty = isDirty;
            IsSupportedByOfficeIMO = isSupportedByOfficeIMO;
            UnsupportedReason = unsupportedReason;
            Dependencies = dependencies ?? Array.Empty<string>();
            DependencyIssues = dependencyIssues ?? Array.Empty<string>();
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 cell reference.</summary>
        public string CellReference { get; }

        /// <summary>Formula text without forcing a leading equals sign.</summary>
        public string Formula { get; }

        /// <summary>Cached cell value, if present.</summary>
        public string? CachedValue { get; }

        /// <summary>True when a cached result is present.</summary>
        public bool HasCachedValue => CachedValue != null;

        /// <summary>True when the formula is marked for recalculation.</summary>
        public bool IsDirty { get; }

        /// <summary>True when OfficeIMO's lightweight evaluator can currently calculate this formula.</summary>
        public bool IsSupportedByOfficeIMO { get; }

        /// <summary>Reason a formula is not supported by OfficeIMO's lightweight evaluator.</summary>
        public string? UnsupportedReason { get; }

        /// <summary>Direct A1/range dependencies detected in the formula.</summary>
        public IReadOnlyList<string> Dependencies { get; }

        /// <summary>Dependency diagnostics such as circular references or missing cached formula dependencies.</summary>
        public IReadOnlyList<string> DependencyIssues { get; }

        /// <summary>True when dependency diagnostics found one or more issues.</summary>
        public bool HasDependencyIssues => DependencyIssues.Count > 0;
    }

    /// <summary>
    /// Describes the current lightweight formula calculation support in OfficeIMO.Excel.
    /// </summary>
    public sealed class ExcelFormulaCapabilities {
        private static readonly string[] Functions = { "SUM", "AVERAGE", "AVERAGEA", "MIN", "MINA", "MAX", "MAXA", "COUNT", "COUNTA", "COUNTBLANK", "SUBTOTAL", "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS", "SUMIFS", "AVERAGEIFS", "MINIFS", "MAXIFS", "PRODUCT", "MEDIAN", "LARGE", "SMALL", "MODE.SNGL", "MODE", "GEOMEAN", "HARMEAN", "AVEDEV", "DEVSQ", "SUMXMY2", "SUMX2MY2", "SUMX2PY2", "SUMSQ", "SUMPRODUCT", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.INC", "QUARTILE.EXC", "PERCENTRANK.INC", "PERCENTRANK.EXC", "RANK.EQ", "RANK.AVG", "COVAR", "COVARIANCE.P", "COVARIANCE.S", "CORREL", "SLOPE", "INTERCEPT", "RSQ", "FORECAST.LINEAR", "PMT", "PV", "FV", "NPER", "NPV", "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "XMATCH", "CONCAT", "CONCATENATE", "TEXT", "TEXTJOIN", "TEXTBEFORE", "TEXTAFTER", "FORMULATEXT", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER", "SUBSTITUTE", "FIND", "SEARCH", "VALUE", "EXACT", "REPT", "ABS", "SIGN", "ROUND", "ROUNDUP", "ROUNDDOWN", "MROUND", "TRUNC", "INT", "CEILING.MATH", "FLOOR.MATH", "CEILING", "FLOOR", "POWER", "SQRT", "LN", "LOG10", "EXP", "PI", "RADIANS", "DEGREES", "MOD", "ROW", "COLUMN", "ROWS", "COLUMNS", "DATE", "TIME", "DATEVALUE", "TIMEVALUE", "TODAY", "NOW", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "DATEDIF", "YEARFRAC", "EDATE", "EOMONTH", "DAYS", "DAYS360", "WEEKDAY", "WEEKNUM", "ISOWEEKNUM", "NETWORKDAYS", "WORKDAY", "WORKDAY.INTL", "IF", "IFS", "SWITCH", "CHOOSE", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISERR", "ISNA", "ISFORMULA", "AND", "OR", "NOT", "IFERROR", "IFNA" };
        private static readonly HashSet<string> FunctionSet = new HashSet<string>(Functions, StringComparer.OrdinalIgnoreCase);
        private static readonly string[] Operators = { "+", "-", "*", "/", ">", "<", ">=", "<=", "=", "<>" };
        private static readonly string[] OperandKinds = { "number literal", "text literal", "same-sheet A1 cell reference", "same-sheet A1 range for function arguments", "cross-sheet A1 cell/range reference", "A1-backed named range reference", "simple table structured reference", "same-sheet numeric/text comparison for IF/IFS/SWITCH/AND/OR/NOT", "bounded dependency depth and circular-reference graph diagnostics" };

        private ExcelFormulaCapabilities() {
        }

        /// <summary>Current OfficeIMO.Excel lightweight formula capability model.</summary>
        public static ExcelFormulaCapabilities Current { get; } = new ExcelFormulaCapabilities();

        internal static bool IsBuiltInFunction(string name) {
            return FunctionSet.Contains(name);
        }

        /// <summary>Supported aggregate functions.</summary>
        public IReadOnlyList<string> SupportedFunctions => Functions;

        /// <summary>Supported binary arithmetic operators.</summary>
        public IReadOnlyList<string> SupportedOperators => Operators;

        /// <summary>Supported operand kinds.</summary>
        public IReadOnlyList<string> SupportedOperandKinds => OperandKinds;

        /// <summary>Maximum formula length accepted by the lightweight evaluator.</summary>
        public int MaxFormulaLength => 8192;

        /// <summary>Short human-readable summary of the current evaluator scope.</summary>
        public string Summary => "Supports simple same-sheet arithmetic plus SUM/AVERAGE/AVERAGEA/MIN/MINA/MAX/MAXA/COUNT/COUNTA/COUNTBLANK/SUBTOTAL/COUNTIF/SUMIF/AVERAGEIF/COUNTIFS/SUMIFS/AVERAGEIFS/MINIFS/MAXIFS/PRODUCT/MEDIAN/LARGE/SMALL/SUMSQ/SUMPRODUCT, bounded MODE.SNGL/MODE/GEOMEAN/HARMEAN/AVEDEV/DEVSQ/SUMXMY2/SUMX2MY2/SUMX2PY2/STDEV.S/STDEV.P/VAR.S/VAR.P/PERCENTILE.INC/PERCENTILE.EXC/QUARTILE.INC/QUARTILE.EXC/PERCENTRANK.INC/PERCENTRANK.EXC/RANK.EQ/RANK.AVG/COVAR/COVARIANCE.P/COVARIANCE.S/CORREL/SLOPE/INTERCEPT/RSQ/FORECAST.LINEAR statistical report formulas, bounded PMT/PV/FV/NPER/NPV financial report formulas, exact-match VLOOKUP/HLOOKUP/INDEX/MATCH/XMATCH returning numeric or text values where applicable, XLOOKUP if-not-found fallbacks plus forward/reverse exact and bounded next-smaller/next-larger numeric search, MATCH/XMATCH bounded exact and next-smaller/next-larger numeric positions, CONCAT/CONCATENATE/TEXT/TEXTJOIN/TEXTBEFORE/TEXTAFTER/FORMULATEXT/LEFT/RIGHT/MID/LEN/TRIM/UPPER/LOWER/PROPER/SUBSTITUTE/FIND/SEARCH/VALUE/EXACT/REPT text helpers, bounded TEXT number/date/time formats for report labels, ABS/SIGN/ROUND/ROUNDUP/ROUNDDOWN/MROUND/TRUNC/INT/CEILING.MATH/FLOOR.MATH/CEILING/FLOOR/POWER/SQRT/LN/LOG10/EXP/PI/RADIANS/DEGREES/MOD, ROW/COLUMN/ROWS/COLUMNS reference-shape helpers, DATE/TIME/DATEVALUE/TIMEVALUE/TODAY/NOW/YEAR/MONTH/DAY/HOUR/MINUTE/SECOND/DATEDIF/YEARFRAC/EDATE/EOMONTH/DAYS/DAYS360/WEEKDAY/WEEKNUM/ISOWEEKNUM/NETWORKDAYS/WORKDAY/WORKDAY.INTL, IF/IFS/SWITCH/CHOOSE with numeric/text comparison or selector branches, ISBLANK/ISNUMBER/ISTEXT/ISERROR/ISERR/ISNA/ISFORMULA report guards, AND/OR/NOT comparisons, and IFERROR/IFNA fallbacks returning numbers or text over numbers, text literals, A1 cells, A1 ranges, A1-backed named ranges, simple table structured references, cross-sheet references, and nested formulas. Inspection reports direct A1 dependencies, formula-to-formula edges, maximum dependency depth, circular-reference groups, and dependency issues for preflight diagnostics.";
    }
}
