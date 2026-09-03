namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    private static void ApplyTaggedTableEvidence(
        PdfReadDocument document,
        int[] selectedPageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates,
        TaggedStructureGraph? graph,
        int maximumArtifactsPerPage,
        PdfUnderstandingWorkBudget workBudget) {
        if (graph is null) return;

        var selectedPageIndexes = new Dictionary<int, int>(selectedPageNumbers.Length);
        var contentIndexes = new TaggedPageContentIndex?[selectedPageNumbers.Length];
        for (int pageIndex = 0; pageIndex < selectedPageNumbers.Length; pageIndex++) {
            workBudget.Consume();
            selectedPageIndexes.Add(selectedPageNumbers[pageIndex], pageIndex);
        }

        var additionsByPage = new List<PdfUnderstandingTableCandidate>?[pages.Count];
        foreach (PdfStructureElementInfo table in graph.Tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(table.ObjectNumber) ||
                !HasResolvedRole(graph.Tagged, table, "Table") ||
                !TryGetTaggedTableRows(graph, table, workBudget, out List<PdfStructureElementInfo> rows) ||
                rows.Count < 2) continue;

            if (!TryReadTaggedRows(
                document,
                graph,
                rows,
                workBudget,
                out int columnCount,
                out List<TaggedTableRow> taggedRows) ||
                columnCount < 2) continue;

            foreach (IGrouping<int, TaggedTableRow> pageRows in taggedRows.GroupBy(static row => row.PageNumber)) {
                workBudget.Consume();
                if (!selectedPageIndexes.TryGetValue(pageRows.Key, out int pageIndex)) continue;
                TaggedTableRow[] sourceRows = pageRows.ToArray();
                workBudget.Consume(sourceRows.Length);
                if (sourceRows.Length < 2) continue;

                TaggedPageContentIndex index = contentIndexes[pageIndex] ??=
                    new TaggedPageContentIndex(
                        document.Pages[pageRows.Key - 1],
                        pages[pageIndex],
                        workBudget);
                PdfUnderstandingTableCandidate? candidate = BuildTaggedTableCandidate(
                    sourceRows,
                    columnCount,
                    index,
                    workBudget);
                if (candidate is null) continue;

                List<PdfUnderstandingTableCandidate> additions = additionsByPage[pageIndex] ??=
                    new List<PdfUnderstandingTableCandidate>();
                additions.Add(candidate);
                if (additions.Count > maximumArtifactsPerPage) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.UnderstandingArtifacts,
                        maximumArtifactsPerPage,
                        additions.Count);
                }
            }
        }

        for (int pageIndex = 0; pageIndex < additionsByPage.Length; pageIndex++) {
            workBudget.Consume();
            List<PdfUnderstandingTableCandidate>? additions = additionsByPage[pageIndex];
            if (additions is null || additions.Count == 0) continue;
            TaggedPageContentIndex index = contentIndexes[pageIndex] ??=
                new TaggedPageContentIndex(
                    document.Pages[selectedPageNumbers[pageIndex] - 1],
                    pages[pageIndex],
                    workBudget);
            HashSet<PdfTextSpan> figureRuns = CollectTaggedFigureRuns(
                document,
                graph,
                selectedPageNumbers[pageIndex],
                index,
                workBudget);
            IReadOnlyList<PdfUnderstandingTableCandidate> existing = RemoveFullyAccountedGeometricTables(
                tableCandidates[pageIndex],
                additions,
                figureRuns,
                workBudget);
            IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
                PdfUnderstandingTableCandidateReconciler.Reconcile(
                    document.Pages[selectedPageNumbers[pageIndex] - 1],
                    existing,
                    additions);
            if (reconciled.Count > maximumArtifactsPerPage) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.UnderstandingArtifacts,
                    maximumArtifactsPerPage,
                    reconciled.Count);
            }
            tableCandidates[pageIndex] = reconciled;
        }
    }

    private static HashSet<PdfTextSpan> CollectTaggedFigureRuns(
        PdfReadDocument document,
        TaggedStructureGraph graph,
        int pageNumber,
        TaggedPageContentIndex contentIndex,
        PdfUnderstandingWorkBudget workBudget) {
        var result = new HashSet<PdfTextSpan>();
        foreach (PdfStructureElementInfo figure in graph.Tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(figure.ObjectNumber) ||
                !HasResolvedRole(graph.Tagged, figure, "Figure") ||
                !TryCollectStructureMarkedContent(
                    document,
                    graph,
                    figure,
                    excludeCaptionSubtrees: true,
                    workBudget,
                    out PageMarkedContent[] content)) continue;

            for (int contentIndexValue = 0; contentIndexValue < content.Length; contentIndexValue++) {
                workBudget.Consume();
                PageMarkedContent item = content[contentIndexValue];
                if (item.PageNumber != pageNumber ||
                    !contentIndex.TryResolveRuns(item.Key, out IReadOnlyList<PdfTextSpan> runs)) continue;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    workBudget.Consume();
                    result.Add(runs[runIndex]);
                }
            }
        }
        return result;
    }

    private static IReadOnlyList<PdfUnderstandingTableCandidate> RemoveFullyAccountedGeometricTables(
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        List<PdfUnderstandingTableCandidate> taggedTables,
        HashSet<PdfTextSpan> figureRuns,
        PdfUnderstandingWorkBudget workBudget) {
        var taggedRunSets = new HashSet<PdfTextSpan>[taggedTables.Count];
        for (int index = 0; index < taggedTables.Count; index++) {
            workBudget.Consume();
            taggedRunSets[index] = PdfUnderstandingTableCandidateReconciler.GetNativeSourceRuns(taggedTables[index]);
        }

        var result = new List<PdfUnderstandingTableCandidate>(existing.Count);
        for (int candidateIndex = 0; candidateIndex < existing.Count; candidateIndex++) {
            workBudget.Consume();
            PdfUnderstandingTableCandidate candidate = existing[candidateIndex];
            HashSet<PdfTextSpan> candidateRuns =
                PdfUnderstandingTableCandidateReconciler.GetNativeSourceRuns(candidate);
            if (candidateRuns.Count == 0) {
                result.Add(candidate);
                continue;
            }

            bool fullyOwnedByNonTableContent = candidateRuns.All(run =>
                run.IsArtifactContent || figureRuns.Contains(run));
            bool fullyAccountedByTableAndFigure = false;
            for (int tableIndex = 0; tableIndex < taggedRunSets.Length && !fullyAccountedByTableAndFigure; tableIndex++) {
                workBudget.Consume();
                HashSet<PdfTextSpan> tableRuns = taggedRunSets[tableIndex];
                fullyAccountedByTableAndFigure =
                    tableRuns.Count > 0 &&
                    candidateRuns.Any(tableRuns.Contains) &&
                    candidateRuns.All(run =>
                        tableRuns.Contains(run) ||
                        run.IsArtifactContent ||
                        figureRuns.Contains(run));
            }
            if (!fullyOwnedByNonTableContent && !fullyAccountedByTableAndFigure) result.Add(candidate);
        }
        return result.Count == 0
            ? Array.Empty<PdfUnderstandingTableCandidate>()
            : result.AsReadOnly();
    }

    private static bool TryGetTaggedTableRows(
        TaggedStructureGraph graph,
        PdfStructureElementInfo table,
        PdfUnderstandingWorkBudget workBudget,
        out List<PdfStructureElementInfo> rows) {
        rows = new List<PdfStructureElementInfo>();
        for (int childIndex = 0; childIndex < table.ChildElementObjectNumbers.Count; childIndex++) {
            workBudget.Consume();
            if (!TryGetReciprocalChild(
                graph,
                table,
                table.ChildElementObjectNumbers[childIndex],
                out PdfStructureElementInfo? child)) return false;
            if (HasResolvedRole(graph.Tagged, child!, "TR")) {
                rows.Add(child!);
                continue;
            }
            if (!HasResolvedRole(graph.Tagged, child!, "THead") &&
                !HasResolvedRole(graph.Tagged, child!, "TBody") &&
                !HasResolvedRole(graph.Tagged, child!, "TFoot")) return false;
            for (int rowIndex = 0; rowIndex < child!.ChildElementObjectNumbers.Count; rowIndex++) {
                workBudget.Consume();
                if (!TryGetReciprocalChild(
                    graph,
                    child,
                    child.ChildElementObjectNumbers[rowIndex],
                    out PdfStructureElementInfo? row) ||
                    !HasResolvedRole(graph.Tagged, row!, "TR")) return false;
                rows.Add(row!);
            }
        }
        return rows.Count > 0;
    }

    private static bool TryReadTaggedRows(
        PdfReadDocument document,
        TaggedStructureGraph graph,
        List<PdfStructureElementInfo> rows,
        PdfUnderstandingWorkBudget workBudget,
        out int columnCount,
        out List<TaggedTableRow> taggedRows) {
        columnCount = 0;
        taggedRows = new List<TaggedTableRow>(rows.Count);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            PdfStructureElementInfo row = rows[rowIndex];
            var cells = new List<PdfStructureElementInfo>(row.ChildElementObjectNumbers.Count);
            for (int childIndex = 0; childIndex < row.ChildElementObjectNumbers.Count; childIndex++) {
                workBudget.Consume();
                if (!TryGetReciprocalChild(
                    graph,
                    row,
                    row.ChildElementObjectNumbers[childIndex],
                    out PdfStructureElementInfo? cell) ||
                    (!HasResolvedRole(graph.Tagged, cell!, "TH") &&
                     !HasResolvedRole(graph.Tagged, cell!, "TD"))) return false;
                cells.Add(cell!);
            }
            if (cells.Count < 2) return false;
            if (columnCount == 0) columnCount = cells.Count;
            else if (cells.Count != columnCount) return false;

            var cellContent = new PageMarkedContent[cells.Count][];
            var rowPages = new HashSet<int>();
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                if (!TryCollectStructureMarkedContent(
                    document,
                    graph,
                    cells[cellIndex],
                    excludeCaptionSubtrees: false,
                    workBudget,
                    out PageMarkedContent[] content)) return false;
                cellContent[cellIndex] = content;
                for (int contentIndex = 0; contentIndex < content.Length; contentIndex++) {
                    workBudget.Consume();
                    rowPages.Add(content[contentIndex].PageNumber);
                }
            }
            if (rowPages.Count != 1) return false;
            int pageNumber = rowPages.Single();
            var keys = new IReadOnlyList<MarkedContentKey>[cells.Count];
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                MarkedContentKey[] pageKeys = cellContent[cellIndex]
                    .Where(content => content.PageNumber == pageNumber)
                    .Select(static content => content.Key)
                    .Distinct()
                    .ToArray();
                workBudget.Consume(pageKeys.Length + 1L);
                keys[cellIndex] = Array.AsReadOnly(pageKeys);
            }
            taggedRows.Add(new TaggedTableRow(pageNumber, keys));
        }
        return taggedRows.Count > 0;
    }

    private static bool TryGetReciprocalChild(
        TaggedStructureGraph graph,
        PdfStructureElementInfo parent,
        int childObjectNumber,
        out PdfStructureElementInfo? child) {
        child = null;
        return graph.ReachableObjectNumbers.Contains(childObjectNumber) &&
            graph.StructuresByObject.TryGetValue(childObjectNumber, out child) &&
            child.ParentObjectNumber == parent.ObjectNumber;
    }

    private static PdfUnderstandingTableCandidate? BuildTaggedTableCandidate(
        TaggedTableRow[] sourceRows,
        int columnCount,
        TaggedPageContentIndex contentIndex,
        PdfUnderstandingWorkBudget workBudget) {
        var rows = new IReadOnlyList<string>[sourceRows.Length];
        var sourceLines = new List<PdfUnderstandingLine>();
        var columnLeft = Enumerable.Repeat(double.PositiveInfinity, columnCount).ToArray();
        var columnRight = Enumerable.Repeat(double.NegativeInfinity, columnCount).ToArray();
        double yTop = double.NegativeInfinity;
        double yBottom = double.PositiveInfinity;

        for (int rowIndex = 0; rowIndex < sourceRows.Length; rowIndex++) {
            var cells = new string[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                workBudget.Consume();
                IReadOnlyList<MarkedContentKey> keys = sourceRows[rowIndex].Cells[columnIndex];
                if (keys.Count == 0) {
                    cells[columnIndex] = string.Empty;
                    continue;
                }
                if (!contentIndex.TryBuildCell(keys, out TaggedTableCell cell)) return null;
                cells[columnIndex] = cell.Text;
                sourceLines.AddRange(cell.SourceLines);
                columnLeft[columnIndex] = Math.Min(columnLeft[columnIndex], cell.Left);
                columnRight[columnIndex] = Math.Max(columnRight[columnIndex], cell.Right);
                yTop = Math.Max(yTop, cell.Top);
                yBottom = Math.Min(yBottom, cell.Bottom);
            }
            rows[rowIndex] = Array.AsReadOnly(cells);
        }

        if (sourceLines.Count == 0 ||
            !IsFinite(yTop) ||
            !IsFinite(yBottom) ||
            columnLeft.Any(double.IsPositiveInfinity) ||
            columnRight.Any(double.IsNegativeInfinity)) return null;

        var columns = new PdfUnderstandingTableColumn[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            workBudget.Consume();
            columns[columnIndex] = new PdfUnderstandingTableColumn(
                columnLeft[columnIndex],
                columnRight[columnIndex]);
        }
        PdfUnderstandingLine[] orderedSourceLines = sourceLines
            .OrderByDescending(static line => line.BaselineY)
            .ThenBy(static line => line.XStart)
            .ToArray();
        workBudget.Consume(orderedSourceLines.Length);
        return new PdfUnderstandingTableCandidate(
            "tagged-structure",
            yTop,
            yBottom,
            Array.AsReadOnly(columns),
            Array.AsReadOnly(rows),
            Array.AsReadOnly(orderedSourceLines),
            0.99D,
            new[] {
                new PdfInferenceEvidence(
                    "table.tagged-structure",
                    "A validated reachable Table, row, and cell hierarchy owns every recovered marked-content cell.",
                    0.99D)
            });
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class TaggedPageContentIndex {
        private readonly PdfReadPage _readPage;
        private readonly PdfUnderstandingPageResult _page;
        private readonly PdfUnderstandingWorkBudget _workBudget;
        private readonly Dictionary<MarkedContentKey, List<PdfTextSpan>> _runsByKey = new();

        internal TaggedPageContentIndex(
            PdfReadPage readPage,
            PdfUnderstandingPageResult page,
            PdfUnderstandingWorkBudget workBudget) {
            _readPage = readPage;
            _page = page;
            _workBudget = workBudget;
            for (int runIndex = 0; runIndex < page.DecodedRuns.Count; runIndex++) {
                workBudget.Consume();
                PdfTextSpan run = page.DecodedRuns[runIndex];
                if (!run.MarkedContentId.HasValue) continue;
                var key = new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value);
                if (!_runsByKey.TryGetValue(key, out List<PdfTextSpan>? runs)) {
                    runs = new List<PdfTextSpan>();
                    _runsByKey.Add(key, runs);
                }
                runs.Add(run);
            }
        }

        internal bool TryBuildCell(
            IReadOnlyList<MarkedContentKey> keys,
            out TaggedTableCell cell) {
            var selectedRuns = new HashSet<PdfTextSpan>();
            for (int keyIndex = 0; keyIndex < keys.Count; keyIndex++) {
                _workBudget.Consume();
                if (!TryResolveRuns(keys[keyIndex], out IReadOnlyList<PdfTextSpan>? runs)) {
                    cell = default;
                    return false;
                }
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    _workBudget.Consume();
                    selectedRuns.Add(runs[runIndex]);
                }
            }
            PdfTextSpan[] orderedRuns = selectedRuns
                .OrderByDescending(static run => run.Y)
                .ThenBy(static run => run.X)
                .ToArray();
            _workBudget.Consume(orderedRuns.Length);
            if (orderedRuns.Length == 0) {
                cell = default;
                return false;
            }

            List<TextLayoutEngine.TextLine> layoutLines = TextLayoutEngine.BuildLines(
                orderedRuns,
                new TextLayoutEngine.Options { ForceSingleColumn = true },
                _workBudget.Consume,
                _workBudget.ThrowIfCancellationRequested);
            string text = string.Join(" ", layoutLines.Select(static line => line.Text)).Trim();
            if (text.Length == 0) {
                cell = default;
                return false;
            }

            var sourceLines = new List<PdfUnderstandingLine>();
            for (int lineIndex = 0; lineIndex < _page.Lines.Count; lineIndex++) {
                _workBudget.Consume();
                PdfUnderstandingLine line = _page.Lines[lineIndex];
                PdfUnderstandingWord[] words = line.Words
                    .Where(word => word.SourceRuns.Any(selectedRuns.Contains))
                    .ToArray();
                if (words.Length == 0) continue;
                _workBudget.Consume(words.Length);
                double tolerance = Math.Max(2.5D, line.FontSize * 0.35D);
                PdfTextSpan[] lineRuns = orderedRuns
                    .Where(run => Math.Abs(run.Y - line.BaselineY) <= tolerance)
                    .ToArray();
                List<TextLayoutEngine.TextLine> cellLines = TextLayoutEngine.BuildLines(
                    lineRuns,
                    new TextLayoutEngine.Options { ForceSingleColumn = true },
                    _workBudget.Consume,
                    _workBudget.ThrowIfCancellationRequested);
                string lineText = string.Join(" ", cellLines.Select(static item => item.Text)).Trim();
                if (lineText.Length == 0) continue;
                sourceLines.Add(new PdfUnderstandingLine(
                    Array.AsReadOnly(words),
                    lineText,
                    words.Average(static word => word.Confidence),
                    line.Evidence,
                    line.SourceKind,
                    line.SourceSequence,
                    line.BlockId,
                    line.ParagraphId,
                    line.LineId));
            }
            if (sourceLines.Count == 0) {
                cell = default;
                return false;
            }

            cell = new TaggedTableCell(
                text,
                orderedRuns.Min(static run => run.X),
                orderedRuns.Max(static run => run.X + Math.Max(0D, run.Advance)),
                orderedRuns.Max(static run => run.Y),
                orderedRuns.Min(static run => run.Y),
                sourceLines.AsReadOnly());
            return true;
        }

        internal bool TryResolveRuns(
            MarkedContentKey key,
            out IReadOnlyList<PdfTextSpan> runs) {
            if (_runsByKey.TryGetValue(key, out List<PdfTextSpan>? exact)) {
                runs = exact;
                return exact.Count > 0;
            }

            List<PdfTextSpan>? match = null;
            foreach (KeyValuePair<MarkedContentKey, List<PdfTextSpan>> candidate in _runsByKey) {
                _workBudget.Consume();
                if (candidate.Key.MarkedContentId != key.MarkedContentId) continue;
                bool compatible = key.ContentStreamObjectNumber.HasValue
                    ? candidate.Key.ContentStreamObjectNumber is null &&
                      _readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)
                    : _readPage.IsPageContentStreamObjectNumber(candidate.Key.ContentStreamObjectNumber);
                if (!compatible) continue;
                if (match is not null) {
                    runs = Array.Empty<PdfTextSpan>();
                    return false;
                }
                match = candidate.Value;
            }
            runs = match is null ? Array.Empty<PdfTextSpan>() : match;
            return runs.Count > 0;
        }
    }

    private readonly record struct TaggedTableRow(
        int PageNumber,
        IReadOnlyList<MarkedContentKey>[] Cells);

    private readonly record struct TaggedTableCell(
        string Text,
        double Left,
        double Right,
        double Top,
        double Bottom,
        IReadOnlyList<PdfUnderstandingLine> SourceLines);
}
