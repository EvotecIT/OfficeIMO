namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    /// <summary>Upper bound on grid columns one tagged table may declare through cells and spans.</summary>
    private const int MaximumTaggedTableColumns = 1024;

    private static void ApplyTaggedTableEvidence(
        PdfReadDocument document,
        int[] selectedPageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates,
        TaggedStructureGraph? graph,
        TaggedContentRoleIndex? taggedRoles,
        int maximumArtifactsPerPage,
        List<PdfTaggedTableDiagnostic>?[] diagnostics,
        PdfUnderstandingWorkBudget workBudget) {
        // Only a catalog that declares itself tagged (/MarkInfo /Marked true) is trusted for table
        // structure; untagged and unmarked files keep geometry without walking table elements.
        if (graph is null || graph.Tagged.Marked != true) return;

        var selectedPageIndexes = new Dictionary<int, List<int>>(selectedPageNumbers.Length);
        var contentIndexes = new TaggedPageContentIndex?[selectedPageNumbers.Length];
        for (int pageIndex = 0; pageIndex < selectedPageNumbers.Length; pageIndex++) {
            workBudget.Consume();
            int pageNumber = selectedPageNumbers[pageIndex];
            if (!selectedPageIndexes.TryGetValue(pageNumber, out List<int>? pageIndexes)) {
                pageIndexes = new List<int>(1);
                selectedPageIndexes.Add(pageNumber, pageIndexes);
            }
            pageIndexes.Add(pageIndex);
        }

        var additionsByPage = new List<PdfUnderstandingTableCandidate>?[pages.Count];
        foreach (PdfStructureElementInfo table in graph.Tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(table.ObjectNumber) ||
                !HasResolvedRole(graph.Tagged, table, "Table")) continue;

            TaggedTableFailure? failure = TryReadTaggedTable(
                document,
                graph,
                table,
                selectedPageIndexes,
                pages,
                contentIndexes,
                workBudget,
                out List<PdfUnderstandingTableCandidate>? candidates,
                out List<int> touchedPages);
            if (failure is TaggedTableFailure reported) {
                AddTaggedTableDiagnostic(document, table, reported, touchedPages, selectedPageIndexes, diagnostics, workBudget);
                continue;
            }

            foreach (PdfUnderstandingTableCandidate candidate in candidates!) {
                workBudget.Consume();
                int pageNumber = candidate.TaggedSegment!.PageNumber;
                List<int> occurrences = selectedPageIndexes[pageNumber];
                for (int occurrence = 0; occurrence < occurrences.Count; occurrence++) {
                    workBudget.Consume();
                    int pageIndex = occurrences[occurrence];
                    List<PdfUnderstandingTableCandidate> additions = additionsByPage[pageIndex] ??=
                        new List<PdfUnderstandingTableCandidate>();
                    additions.Add(occurrence == 0 ? candidate : candidate.CopyForOccurrence());
                    if (additions.Count > maximumArtifactsPerPage) {
                        throw PdfReadLimitException.Create(
                            PdfReadLimitKind.UnderstandingArtifacts,
                            maximumArtifactsPerPage,
                            additions.Count);
                    }
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
            int pageNumber = selectedPageNumbers[pageIndex];
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            IReadOnlyList<PdfUnderstandingTableCandidate> existing = RemoveFullyAccountedGeometricTables(
                tableCandidates[pageIndex],
                additions,
                figureRuns,
                workBudget,
                run => IsTaggedProseOutsideTables(run, additions, taggedRoles, pageNumber, readPage, workBudget));
            IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
                PdfUnderstandingTableCandidateReconciler.Reconcile(
                    document.Pages[selectedPageNumbers[pageIndex] - 1],
                    existing,
                    additions,
                    workBudget.Consume,
                    workBudget.ThrowIfCancellationRequested);
            if (reconciled.Count > maximumArtifactsPerPage) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.UnderstandingArtifacts,
                    maximumArtifactsPerPage,
                    reconciled.Count);
            }
            tableCandidates[pageIndex] = reconciled;
        }
    }

    /// <summary>
    /// Reads one tagged table into page segments. Returns null on success; otherwise the reason
    /// the table falls back to geometric detection.
    /// </summary>
    private static TaggedTableFailure? TryReadTaggedTable(
        PdfReadDocument document,
        TaggedStructureGraph graph,
        PdfStructureElementInfo table,
        Dictionary<int, List<int>> selectedPageIndexes,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        TaggedPageContentIndex?[] contentIndexes,
        PdfUnderstandingWorkBudget workBudget,
        out List<PdfUnderstandingTableCandidate>? candidates,
        out List<int> touchedPages) {
        candidates = null;
        touchedPages = new List<int>();
        if (!TryGetTaggedTableRows(graph, table, workBudget, out List<TaggedTableRowSource> rows)) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnsupportedStructure,
                "The Table element contains children other than THead, TBody, TFoot, TR, TH and TD, or a child that does not name its parent.");
        }

        TaggedTableFailure? gridFailure = TryPlaceTaggedCells(rows, workBudget, out List<TaggedGridRow> grid, out int columnCount);
        if (gridFailure is not null) return gridFailure;

        // Body rows take their content from one page each and must not move back through pages.
        // A tagged header row may also own the header repeated at the top of later pages; each
        // page's copy becomes that page's first row.
        var pagedRows = new List<TaggedGridRow>(grid.Count);
        var pendingHeaders = new List<TaggedGridRow>();
        int previousPage = 0;
        for (int rowIndex = 0; rowIndex < grid.Count; rowIndex++) {
            TaggedGridRow row = grid[rowIndex];
            var rowPages = new List<int>(1);
            foreach (TaggedGridCell cell in row.Cells) {
                if (!TryCollectStructureMarkedContent(document, graph, cell.Element, excludeCaptionSubtrees: false, workBudget, out PageMarkedContent[] content)) {
                    return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnsupportedStructure,
                        "A table cell has a descendant that is unreachable or does not name its parent.");
                }
                cell.Content = content;
                for (int contentIndex = 0; contentIndex < content.Length; contentIndex++) {
                    workBudget.Consume();
                    int page = content[contentIndex].PageNumber;
                    if (!touchedPages.Contains(page)) touchedPages.Add(page);
                    if (!rowPages.Contains(page)) rowPages.Add(page);
                }
            }
            rowPages.Sort();
            if (rowPages.Count > 1 && !row.IsHeader) {
                return new TaggedTableFailure(PdfTaggedTableFallbackReason.RowSpansPages,
                    "Row " + (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " has cell content on pages " + rowPages[0].ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " and " + rowPages[1].ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }
            if (rowPages.Count == 0) {
                row.PageNumber = 0;
                pagedRows.Add(row);
                continue;
            }
            if (rowPages[0] < previousPage) {
                return new TaggedTableFailure(PdfTaggedTableFallbackReason.RowSpansPages,
                    "Row " + (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " is on an earlier page than the row before it.");
            }
            row.PageNumber = rowPages[0];
            pagedRows.Add(row);
            previousPage = rowPages[0];
            for (int pageIndex = 1; pageIndex < rowPages.Count; pageIndex++) {
                workBudget.Consume();
                pendingHeaders.Add(row.CopyForPage(rowPages[pageIndex]));
            }
        }
        if (pendingHeaders.Count > 0) {
            // Place each repeated header before the first body row of its page.
            var ordered = new List<TaggedGridRow>(pagedRows.Count + pendingHeaders.Count);
            int lastPage = 0;
            foreach (TaggedGridRow row in pagedRows) {
                workBudget.Consume();
                if (row.PageNumber > lastPage) {
                    foreach (TaggedGridRow repeated in pendingHeaders) {
                        if (repeated.PageNumber > lastPage && repeated.PageNumber <= row.PageNumber) ordered.Add(repeated);
                    }
                    lastPage = row.PageNumber;
                }
                ordered.Add(row);
            }
            foreach (TaggedGridRow repeated in pendingHeaders) {
                if (repeated.PageNumber > lastPage) ordered.Add(repeated);
            }
            grid = ordered;
        } else {
            grid = pagedRows;
        }

        // Resolve each selected page's cell text from the decoded runs that carry the referenced MCIDs.
        bool anySelected = false;
        foreach (TaggedGridRow row in grid) {
            workBudget.Consume();
            if (row.PageNumber == 0 || !selectedPageIndexes.TryGetValue(row.PageNumber, out List<int>? pageIndexes)) continue;
            anySelected = true;
            int pageIndex = pageIndexes[0];
            TaggedPageContentIndex index = contentIndexes[pageIndex] ??=
                new TaggedPageContentIndex(document.Pages[row.PageNumber - 1], pages[pageIndex], workBudget);
            foreach (TaggedGridCell cell in row.Cells) {
                MarkedContentKey[] keys = cell.ContentOn(row.PageNumber)
                    .Distinct()
                    .ToArray();
                workBudget.Consume(keys.Length + 1L);
                if (!index.TryBuildCell(keys, out TaggedTableCell? resolved)) {
                    return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnresolvedContent,
                        "A cell refers to marked content that page " + row.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        " does not contain as text.");
                }
                cell.Resolved = resolved;
            }
        }
        if (!anySelected) return null;
        if (!HasConsistentCellOrder(grid, workBudget, out int conflictingRow)) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.CellOrderConflictsWithLayout,
                "Row " + (conflictingRow + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " lists its cells in a different order than they appear on the page.");
        }

        // Drop rows and columns that hold no text anywhere in the selected rows: spreadsheet exports
        // tag unused grid tracks and spacer rows. Empty cells inside a kept row and column remain.
        var keptRows = grid.Where(static row => row.PageNumber != 0 && row.Cells.Any(static cell => cell.Resolved is not null)).ToList();
        var keptColumn = new bool[columnCount];
        foreach (TaggedGridRow row in keptRows) {
            foreach (TaggedGridCell cell in row.Cells) {
                workBudget.Consume();
                if (cell.Resolved is not null) keptColumn[cell.Column] = true;
            }
        }
        var columnMap = new int[columnCount];
        int keptColumnCount = 0;
        for (int column = 0; column < columnCount; column++) {
            columnMap[column] = keptColumn[column] ? keptColumnCount++ : -1;
        }
        if (keptRows.Count == 0) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.NoContent, "The tagged table holds no text.");
        }
        if (keptColumnCount < 2 || keptRows.Count < 2 ||
            !keptRows.Any(static row => row.Cells.Count(static cell => cell.Resolved is not null) >= 2)) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.LayoutOnly,
                "The tagged table has fewer than two rows with text or no row with two filled cells, so it is read as page layout.");
        }

        IReadOnlyList<string>? header = keptRows[0].IsHeader
            ? BuildTaggedRowTexts(keptRows[0], columnMap, keptColumnCount)
            : null;

        candidates = new List<PdfUnderstandingTableCandidate>();
        int segmentIndex = 0;
        for (int start = 0; start < keptRows.Count;) {
            int pageNumber = keptRows[start].PageNumber;
            int end = start;
            while (end < keptRows.Count && keptRows[end].PageNumber == pageNumber) end++;
            if (selectedPageIndexes.ContainsKey(pageNumber)) {
                PdfUnderstandingTableCandidate? candidate = BuildTaggedTableCandidate(
                    table.ObjectNumber,
                    pageNumber,
                    segmentIndex,
                    keptRows.GetRange(start, end - start),
                    columnMap,
                    keptColumnCount,
                    keptRows[start].IsHeader,
                    keptRows[start].IsHeader ? null : header,
                    workBudget);
                if (candidate is null) {
                    candidates = null;
                    return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnresolvedContent,
                        "Cells on page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        " could not be placed on the page's text lines.");
                }
                candidates.Add(candidate);
            }
            segmentIndex++;
            start = end;
        }
        return null;
    }

    /// <summary>
    /// Checks that tagged cell order runs the same horizontal direction as the cells on the page in
    /// every row. Either direction is accepted, so right-to-left tables pass; rotated text is not judged.
    /// </summary>
    private static bool HasConsistentCellOrder(List<TaggedGridRow> grid, PdfUnderstandingWorkBudget workBudget, out int conflictingRow) {
        const double Tolerance = 1D;
        int direction = 0;
        conflictingRow = -1;
        for (int rowIndex = 0; rowIndex < grid.Count; rowIndex++) {
            TaggedTableCell? previous = null;
            foreach (TaggedGridCell cell in grid[rowIndex].Cells) {
                workBudget.Consume();
                if (cell.Resolved is not TaggedTableCell current) continue;
                if (current.SourceRuns.Any(static run => Math.Abs(run.RotationDegrees) > 0.5D)) return true;
                if (previous is TaggedTableCell before) {
                    double step = current.Left - before.Left;
                    int stepDirection = step > Tolerance ? 1 : step < -Tolerance ? -1 : 0;
                    if (stepDirection != 0) {
                        if (direction == 0) direction = stepDirection;
                        else if (stepDirection != direction) {
                            conflictingRow = grid[rowIndex].GridIndex;
                            return false;
                        }
                    }
                }
                previous = current;
            }
        }
        return true;
    }

    private static string[] BuildTaggedRowTexts(TaggedGridRow row, int[] columnMap, int columnCount) {
        var cells = new string[columnCount];
        for (int index = 0; index < cells.Length; index++) cells[index] = string.Empty;
        foreach (TaggedGridCell cell in row.Cells) {
            int column = columnMap[cell.Column];
            if (column >= 0 && cell.Resolved is TaggedTableCell resolved) cells[column] = resolved.Text;
        }
        return cells;
    }

    /// <summary>
    /// Places TH/TD cells on a grid, honouring RowSpan and ColSpan. Positions a spanning cell covers stay empty.
    /// </summary>
    private static TaggedTableFailure? TryPlaceTaggedCells(
        List<TaggedTableRowSource> rows,
        PdfUnderstandingWorkBudget workBudget,
        out List<TaggedGridRow> grid,
        out int columnCount) {
        grid = new List<TaggedGridRow>(rows.Count);
        columnCount = 0;
        // Rows still covered by a cell above, per column.
        var covered = new List<int>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            TaggedTableRowSource source = rows[rowIndex];
            var row = new TaggedGridRow(source.IsHeader, rowIndex);
            bool[] coveredAtStart = covered.Select(static remaining => remaining > 0).ToArray();
            int column = 0;
            for (int cellIndex = 0; cellIndex < source.Cells.Count; cellIndex++) {
                workBudget.Consume();
                PdfStructureElementInfo element = source.Cells[cellIndex];
                while (column < coveredAtStart.Length && coveredAtStart[column]) column++;
                int columnSpan = element.ColumnSpan;
                if ((long)column + columnSpan > MaximumTaggedTableColumns) {
                    return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnsupportedStructure,
                        "The tagged table declares more than " + MaximumTaggedTableColumns.ToString(System.Globalization.CultureInfo.InvariantCulture) + " columns.");
                }
                while (covered.Count < column + columnSpan) covered.Add(0);
                for (int spanned = column; spanned < column + columnSpan; spanned++) {
                    workBudget.Consume();
                    if (spanned < coveredAtStart.Length && coveredAtStart[spanned]) {
                        return new TaggedTableFailure(PdfTaggedTableFallbackReason.UnsupportedStructure,
                            "Cell spans overlap in row " + (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
                    }
                    covered[spanned] = element.RowSpan;
                }
                row.Cells.Add(new TaggedGridCell(element, column, element.RowSpan, columnSpan, source.IsHeader || source.HeaderCells[cellIndex]));
                column += columnSpan;
            }
            columnCount = Math.Max(columnCount, covered.Count);
            for (int index = 0; index < covered.Count; index++) {
                if (covered[index] > 0) covered[index]--;
            }
            if (row.Cells.Count > 0 && !source.IsHeader && row.Cells.All(static cell => cell.IsHeader) && grid.All(static previous => previous.Cells.Count == 0 || previous.IsHeader)) {
                row.IsHeader = true;
            }
            grid.Add(row);
        }
        return null;
    }

    private static void AddTaggedTableDiagnostic(
        PdfReadDocument document,
        PdfStructureElementInfo table,
        TaggedTableFailure failure,
        List<int> touchedPages,
        Dictionary<int, List<int>> selectedPageIndexes,
        List<PdfTaggedTableDiagnostic>?[] diagnostics,
        PdfUnderstandingWorkBudget workBudget) {
        var pagesToReport = new List<int>(touchedPages);
        if (pagesToReport.Count == 0 && table.PageObjectNumber is int pageObject &&
            document.GetPageNumberForObject(pageObject, workBudget.CancellationToken) is int tablePage) {
            pagesToReport.Add(tablePage);
        }
        foreach (int pageNumber in pagesToReport) {
            workBudget.Consume();
            if (!selectedPageIndexes.TryGetValue(pageNumber, out List<int>? pageIndexes)) continue;
            foreach (int pageIndex in pageIndexes) {
                (diagnostics[pageIndex] ??= new List<PdfTaggedTableDiagnostic>()).Add(
                    new PdfTaggedTableDiagnostic(table.ObjectNumber, pageNumber, failure.Reason, failure.Message));
            }
        }
    }

    private static bool TryGetTaggedTableRows(
        TaggedStructureGraph graph,
        PdfStructureElementInfo table,
        PdfUnderstandingWorkBudget workBudget,
        out List<TaggedTableRowSource> rows) {
        rows = new List<TaggedTableRowSource>();
        for (int childIndex = 0; childIndex < table.ChildElementObjectNumbers.Count; childIndex++) {
            workBudget.Consume();
            if (!TryGetReciprocalChild(
                graph,
                table,
                table.ChildElementObjectNumbers[childIndex],
                out PdfStructureElementInfo? child)) return false;
            if (HasResolvedRole(graph.Tagged, child!, "TR")) {
                if (!TryAddTaggedRow(graph, child!, isHeader: false, workBudget, rows)) return false;
                continue;
            }
            if (HasResolvedRole(graph.Tagged, child!, "Caption")) continue;
            bool isHeader = HasResolvedRole(graph.Tagged, child!, "THead");
            if (!isHeader &&
                !HasResolvedRole(graph.Tagged, child!, "TBody") &&
                !HasResolvedRole(graph.Tagged, child!, "TFoot")) return false;
            for (int rowIndex = 0; rowIndex < child!.ChildElementObjectNumbers.Count; rowIndex++) {
                workBudget.Consume();
                if (!TryGetReciprocalChild(
                    graph,
                    child,
                    child.ChildElementObjectNumbers[rowIndex],
                    out PdfStructureElementInfo? row) ||
                    !HasResolvedRole(graph.Tagged, row!, "TR") ||
                    !TryAddTaggedRow(graph, row!, isHeader, workBudget, rows)) return false;
            }
        }
        return rows.Count > 0;
    }

    private static bool TryAddTaggedRow(
        TaggedStructureGraph graph,
        PdfStructureElementInfo row,
        bool isHeader,
        PdfUnderstandingWorkBudget workBudget,
        List<TaggedTableRowSource> rows) {
        var cells = new List<PdfStructureElementInfo>(row.ChildElementObjectNumbers.Count);
        var headerCells = new List<bool>(row.ChildElementObjectNumbers.Count);
        for (int childIndex = 0; childIndex < row.ChildElementObjectNumbers.Count; childIndex++) {
            workBudget.Consume();
            if (!TryGetReciprocalChild(
                graph,
                row,
                row.ChildElementObjectNumbers[childIndex],
                out PdfStructureElementInfo? cell)) return false;
            bool header = HasResolvedRole(graph.Tagged, cell!, "TH");
            if (!header && !HasResolvedRole(graph.Tagged, cell!, "TD")) return false;
            cells.Add(cell!);
            headerCells.Add(header);
        }
        rows.Add(new TaggedTableRowSource(isHeader, cells, headerCells));
        return true;
    }

    private static PdfUnderstandingTableCandidate? BuildTaggedTableCandidate(
        int tableObjectNumber,
        int pageNumber,
        int segmentIndex,
        List<TaggedGridRow> sourceRows,
        int[] columnMap,
        int columnCount,
        bool firstRowIsHeader,
        IReadOnlyList<string>? inheritedHeader,
        PdfUnderstandingWorkBudget workBudget) {
        var rows = new IReadOnlyList<string>[sourceRows.Count];
        var sourceLines = new List<PdfUnderstandingLine>();
        var sourceRuns = new HashSet<PdfTextSpan>();
        var spans = new List<PdfTaggedTableCellSpan>();
        var columnLeft = Enumerable.Repeat(double.PositiveInfinity, columnCount).ToArray();
        var columnRight = Enumerable.Repeat(double.NegativeInfinity, columnCount).ToArray();
        var spanningLeft = Enumerable.Repeat(double.PositiveInfinity, columnCount).ToArray();
        var spanningRight = Enumerable.Repeat(double.NegativeInfinity, columnCount).ToArray();
        double yTop = double.NegativeInfinity;
        double yBottom = double.PositiveInfinity;
        var rowPositions = new Dictionary<TaggedGridRow, int>(sourceRows.Count);
        for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) rowPositions.Add(sourceRows[rowIndex], rowIndex);

        for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) {
            TaggedGridRow row = sourceRows[rowIndex];
            var cells = new string[columnCount];
            for (int index = 0; index < cells.Length; index++) cells[index] = string.Empty;
            foreach (TaggedGridCell cell in row.Cells) {
                workBudget.Consume();
                int column = columnMap[cell.Column];
                if (column < 0) continue;
                int columnSpan = 0;
                for (int spanned = cell.Column; spanned < cell.Column + cell.ColumnSpan && spanned < columnMap.Length; spanned++) {
                    if (columnMap[spanned] >= 0) columnSpan++;
                }
                int rowSpan = 1;
                for (int later = rowIndex + 1; later < sourceRows.Count && sourceRows[later].GridIndex < (long)row.GridIndex + cell.RowSpan; later++) {
                    rowSpan++;
                }
                if (rowSpan > 1 || columnSpan > 1) spans.Add(new PdfTaggedTableCellSpan(rowIndex, column, rowSpan, columnSpan));
                if (cell.Resolved is not TaggedTableCell resolved) continue;
                cells[column] = resolved.Text;
                sourceLines.AddRange(resolved.SourceLines);
                for (int runIndex = 0; runIndex < resolved.SourceRuns.Count; runIndex++) {
                    workBudget.Consume();
                    sourceRuns.Add(resolved.SourceRuns[runIndex]);
                }
                double[] left = columnSpan == 1 ? columnLeft : spanningLeft;
                double[] right = columnSpan == 1 ? columnRight : spanningRight;
                left[column] = Math.Min(left[column], resolved.Left);
                right[column] = Math.Max(right[column], resolved.Right);
                yTop = Math.Max(yTop, resolved.Top);
                yBottom = Math.Min(yBottom, resolved.Bottom);
            }
            rows[rowIndex] = Array.AsReadOnly(cells);
        }

        if (sourceLines.Count == 0 || !IsFinite(yTop) || !IsFinite(yBottom)) return null;

        // Columns without a single-column cell on this page borrow a spanning cell's
        // extent, then the gap between their neighbours, so every column has an ordered range.
        for (int column = 0; column < columnCount; column++) {
            if (!double.IsPositiveInfinity(columnLeft[column])) continue;
            columnLeft[column] = spanningLeft[column];
            columnRight[column] = spanningRight[column];
        }
        double tableLeft = sourceRuns.Min(static run => run.X);
        double tableRight = sourceRuns.Max(static run => run.X + Math.Max(0D, run.Advance));
        var columns = new PdfUnderstandingTableColumn[columnCount];
        for (int column = 0; column < columnCount; column++) {
            workBudget.Consume();
            double from = columnLeft[column];
            double to = columnRight[column];
            if (double.IsPositiveInfinity(from) || double.IsNegativeInfinity(to)) {
                double previous = tableLeft;
                for (int index = column - 1; index >= 0; index--) {
                    if (!double.IsNegativeInfinity(columnRight[index])) { previous = columnRight[index]; break; }
                }
                double next = tableRight;
                for (int index = column + 1; index < columnCount; index++) {
                    if (!double.IsPositiveInfinity(columnLeft[index])) { next = columnLeft[index]; break; }
                }
                from = previous;
                to = Math.Max(previous, next);
                columnLeft[column] = from;
                columnRight[column] = to;
            }
            columns[column] = new PdfUnderstandingTableColumn(from, to);
        }

        PdfUnderstandingLine[] orderedSourceLines = sourceLines
            .OrderByDescending(static line => line.BaselineY)
            .ThenBy(static line => line.XStart)
            .ToArray();
        workBudget.Consume(orderedSourceLines.Length);
        var evidence = new List<PdfInferenceEvidence> {
            new PdfInferenceEvidence(
                "table-tagged-structure",
                "A validated reachable Table, row, and cell hierarchy owns every recovered marked-content cell.",
                0.99D)
        };
        if (firstRowIsHeader) {
            evidence.Add(new PdfInferenceEvidence(
                "table-tagged-header-row",
                "Tagged-PDF THead or top-row TH structure identifies the first table row as column headers.",
                0.99D));
        }
        PdfUnderstandingTableCandidate candidate = PdfUnderstandingTableCandidate.FromTagged(
            yTop,
            yBottom,
            Array.AsReadOnly(columns),
            Array.AsReadOnly(rows),
            Array.AsReadOnly(orderedSourceLines),
            sourceRuns.ToArray(),
            0.99D,
            evidence,
            workBudget.Consume,
            workBudget.ThrowIfCancellationRequested);
        candidate.TaggedSegment = new PdfTaggedTableSegment(
            tableObjectNumber,
            pageNumber,
            segmentIndex,
            inheritedHeader,
            spans.Count == 0 ? Array.Empty<PdfTaggedTableCellSpan>() : spans.ToArray());
        return candidate;
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

    private static bool IsTaggedProseOutsideTables(
        PdfTextSpan run,
        List<PdfUnderstandingTableCandidate> tables,
        TaggedContentRoleIndex? taggedRoles,
        int pageNumber,
        PdfReadPage readPage,
        PdfUnderstandingWorkBudget workBudget) {
        if (taggedRoles is null || !run.MarkedContentId.HasValue) return false;
        IReadOnlyList<string> roles = taggedRoles.GetRoles(pageNumber, readPage,
            new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value));
        if (roles.Count == 0 || roles.Contains("Table", StringComparer.OrdinalIgnoreCase)) return false;
        // Headings and other tagged blocks outside every Table element belong to the reading
        // flow wherever they sit; prose keeps the stricter outside-the-grid test below.
        if (!roles.Contains("P", StringComparer.OrdinalIgnoreCase) &&
            !roles.Contains("Caption", StringComparer.OrdinalIgnoreCase)) return true;

        foreach (PdfUnderstandingTableCandidate table in tables) {
            workBudget.Consume();
            if (run.Y >= table.YBottom && run.Y <= table.YTop) return false;
        }
        return true;
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

    internal static IReadOnlyList<PdfUnderstandingTableCandidate> RemoveFullyAccountedGeometricTables(
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        List<PdfUnderstandingTableCandidate> taggedTables,
        HashSet<PdfTextSpan> figureRuns,
        PdfUnderstandingWorkBudget workBudget,
        Func<PdfTextSpan, bool>? isTaggedProseOutsideTables = null) {
        var taggedRuns = new HashSet<PdfTextSpan>();
        for (int index = 0; index < taggedTables.Count; index++) {
            workBudget.Consume();
            IReadOnlyList<PdfTextSpan> tableRuns = taggedTables[index].NativeSourceRuns;
            for (int runIndex = 0; runIndex < tableRuns.Count; runIndex++) {
                workBudget.Consume();
                taggedRuns.Add(tableRuns[runIndex]);
            }
        }

        var result = new List<PdfUnderstandingTableCandidate>(existing.Count);
        for (int candidateIndex = 0; candidateIndex < existing.Count; candidateIndex++) {
            workBudget.Consume();
            PdfUnderstandingTableCandidate candidate = existing[candidateIndex];
            IReadOnlyList<PdfTextSpan> candidateRuns = candidate.NativeSourceRuns;
            if (candidateRuns.Count == 0) {
                result.Add(candidate);
                continue;
            }

            bool intersectsTaggedTable = false;
            bool fullyAccounted = true;
            bool fullyOwnedByNonTableContent = true;
            for (int runIndex = 0; runIndex < candidateRuns.Count; runIndex++) {
                workBudget.Consume();
                PdfTextSpan run = candidateRuns[runIndex];
                bool ownedByTaggedTable = taggedRuns.Contains(run);
                bool ownedByNonTableContent = run.IsArtifactContent || figureRuns.Contains(run);
                intersectsTaggedTable |= ownedByTaggedTable;
                fullyOwnedByNonTableContent &= ownedByNonTableContent;
                // A separately tagged paragraph outside the grid belongs in the reading flow.
                // It cannot keep a competing candidate that otherwise repeats authoritative cells.
                fullyAccounted = fullyAccounted && (ownedByTaggedTable || ownedByNonTableContent ||
                    string.IsNullOrWhiteSpace(run.Text) || isTaggedProseOutsideTables?.Invoke(run) == true);
            }
            if (!fullyOwnedByNonTableContent && !(intersectsTaggedTable && fullyAccounted)) result.Add(candidate);
        }
        return result.Count == 0
            ? Array.Empty<PdfUnderstandingTableCandidate>()
            : result.AsReadOnly();
    }


    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class TaggedPageContentIndex {
        private readonly PdfReadPage _readPage;
        private readonly PdfUnderstandingPageResult _page;
        private readonly PdfUnderstandingWorkBudget _workBudget;
        private readonly Dictionary<MarkedContentKey, List<PdfTextSpan>> _runsByKey = new();
        private Dictionary<int, int?>? _textlessPageContent;

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
            out TaggedTableCell? cell) {
            var selectedRuns = new HashSet<PdfTextSpan>();
            for (int keyIndex = 0; keyIndex < keys.Count; keyIndex++) {
                _workBudget.Consume();
                if (!TryResolveRuns(keys[keyIndex], out IReadOnlyList<PdfTextSpan>? runs)) {
                    if (IsTextlessPageContent(keys[keyIndex])) continue;
                    cell = default;
                    return false;
                }
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    _workBudget.Consume();
                    selectedRuns.Add(runs[runIndex]);
                }
            }
            if (selectedRuns.Count == 0) {
                cell = null;
                return true;
            }

            var orderedRuns = new PdfTextSpan[selectedRuns.Count];
            int orderedRunIndex = 0;
            foreach (PdfTextSpan run in selectedRuns) {
                _workBudget.Consume();
                orderedRuns[orderedRunIndex++] = run;
            }
            Array.Sort(orderedRuns, (left, right) => {
                _workBudget.Consume();
                int baseline = right.Y.CompareTo(left.Y);
                return baseline != 0 ? baseline : left.X.CompareTo(right.X);
            });
            List<TextLayoutEngine.TextLine> layoutLines = TextLayoutEngine.BuildLines(
                orderedRuns,
                new TextLayoutEngine.Options { ForceSingleColumn = true },
                _workBudget.Consume,
                _workBudget.ThrowIfCancellationRequested);
            string text = string.Join(" ", layoutLines.Select(static line => line.Text)).Trim();
            if (text.Length == 0) {
                cell = null;
                return orderedRuns.All(static run => string.IsNullOrWhiteSpace(run.Text));
            }

            var sourceLines = new List<PdfUnderstandingLine>();
            for (int lineIndex = 0; lineIndex < _page.Lines.Count; lineIndex++) {
                _workBudget.Consume();
                PdfUnderstandingLine line = _page.Lines[lineIndex];
                double tolerance = Math.Max(2.5D, line.FontSize * 0.35D);
                var matchingRuns = new List<PdfTextSpan>();
                for (int runIndex = 0; runIndex < orderedRuns.Length; runIndex++) {
                    _workBudget.Consume();
                    PdfTextSpan run = orderedRuns[runIndex];
                    if (Math.Abs(run.Y - line.BaselineY) <= tolerance) matchingRuns.Add(run);
                }
                PdfTextSpan[] lineRuns = matchingRuns.ToArray();
                if (lineRuns.Length == 0) continue;
                List<TextLayoutEngine.TextLine> cellLines = TextLayoutEngine.BuildLines(
                    lineRuns,
                    new TextLayoutEngine.Options { ForceSingleColumn = true },
                    _workBudget.Consume,
                    _workBudget.ThrowIfCancellationRequested);
                string lineText = string.Join(" ", cellLines.Select(static item => item.Text)).Trim();
                if (lineText.Length == 0) continue;
                var word = new PdfUnderstandingWord(
                    lineText,
                    lineRuns.Min(static run => run.X),
                    lineRuns.Max(static run => run.X + Math.Max(0D, run.Advance)),
                    lineRuns.Average(static run => run.Y),
                    lineRuns.Max(static run => run.FontSize),
                    lineRuns.Average(static run => run.RotationDegrees),
                    Array.AsReadOnly(lineRuns),
                    line.Confidence,
                    line.Evidence,
                    lineRuns.Sum(static run => Math.Max(0D, run.Advance)),
                    sourceSequence: line.SourceSequence);
                sourceLines.Add(new PdfUnderstandingLine(
                    new[] { word },
                    lineText,
                    line.Confidence,
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
                sourceLines.AsReadOnly(),
                Array.AsReadOnly(orderedRuns));
            return true;
        }

        private bool IsTextlessPageContent(MarkedContentKey key) {
            if (!_readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)) return false;
            _textlessPageContent ??= _readPage.GetTextlessPageMarkedContent(
                () => _workBudget.Consume(), _workBudget.CancellationToken);
            return _textlessPageContent.TryGetValue(key.MarkedContentId, out int? stream) &&
                (!key.ContentStreamObjectNumber.HasValue || key.ContentStreamObjectNumber == stream);
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

    private readonly record struct TaggedTableRowSource(
        bool IsHeader,
        List<PdfStructureElementInfo> Cells,
        List<bool> HeaderCells);

    private readonly record struct TaggedTableFailure(PdfTaggedTableFallbackReason Reason, string Message);

    private sealed class TaggedGridRow {
        internal TaggedGridRow(bool isHeader, int gridIndex) {
            IsHeader = isHeader;
            GridIndex = gridIndex;
        }
        internal bool IsHeader { get; set; }
        internal int GridIndex { get; }
        internal int PageNumber { get; set; }
        internal List<TaggedGridCell> Cells { get; } = new();

        /// <summary>A copy of this header row holding the cells' content on another page.</summary>
        internal TaggedGridRow CopyForPage(int pageNumber) {
            var copy = new TaggedGridRow(IsHeader, GridIndex) { PageNumber = pageNumber };
            foreach (TaggedGridCell cell in Cells) {
                copy.Cells.Add(new TaggedGridCell(cell.Element, cell.Column, cell.RowSpan, cell.ColumnSpan, cell.IsHeader) {
                    Content = cell.Content
                });
            }
            return copy;
        }
    }

    private sealed class TaggedGridCell {
        internal TaggedGridCell(PdfStructureElementInfo element, int column, int rowSpan, int columnSpan, bool isHeader) {
            Element = element;
            Column = column;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
            IsHeader = isHeader;
        }
        internal PdfStructureElementInfo Element { get; }
        internal int Column { get; }
        internal int RowSpan { get; }
        internal int ColumnSpan { get; }
        internal bool IsHeader { get; }
        internal PageMarkedContent[] Content { get; set; } = Array.Empty<PageMarkedContent>();
        internal IEnumerable<MarkedContentKey> ContentOn(int pageNumber) {
            foreach (PageMarkedContent item in Content) {
                if (item.PageNumber == pageNumber) yield return item.Key;
            }
        }
        internal TaggedTableCell? Resolved { get; set; }
    }

    private readonly record struct TaggedTableCell(
        string Text,
        double Left,
        double Right,
        double Top,
        double Bottom,
        IReadOnlyList<PdfUnderstandingLine> SourceLines,
        IReadOnlyList<PdfTextSpan> SourceRuns);
}
