namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    /// <summary>Upper bound on grid columns one tagged table may declare through cells and spans.</summary>
    private const int MaximumTaggedTableColumns = 1024;

    /// <summary>
    /// Reads the structure tree once, before any page is analysed, for the pages its tagged tables name (in a catalog
    /// trusted for table structure, see <see cref="ApplyTaggedTableEvidence"/>). Null when there are none.
    /// </summary>
    internal static TaggedTableReads? PrepareTaggedTables(PdfReadDocument document, long maxWorkUnits, CancellationToken cancellationToken) {
        var workBudget = new PdfUnderstandingWorkBudget(maxWorkUnits, cancellationToken);
        TaggedStructureGraph? graph = BuildTaggedStructureGraph(document, workBudget);
        if (graph is null || graph.Tagged.Marked != true) return null;
        var pages = new HashSet<int>();
        foreach (PdfStructureElementInfo table in graph.Tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(table.ObjectNumber) || !HasResolvedRole(graph.Tagged, table, "Table")) continue;
            foreach (int pageObject in CollectTablePageObjects(graph, table, workBudget)) {
                if (document.GetPageNumberForObject(pageObject, cancellationToken) is int page) pages.Add(page);
            }
        }
        return pages.Count == 0 ? null : new TaggedTableReads(graph, pages);
    }

    /// <summary>
    /// Reads every tagged table from the analysed pages into <paramref name="reads"/>, ahead of enrichment, and says
    /// which pages none of them supplies: a page a failing table names, or one no table passes on. Enrichment then
    /// reuses the reads.
    /// </summary>
    internal static bool[] ReadTaggedTables(
        PdfReadDocument document,
        int[] selectedPageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        TaggedTableReads reads,
        long maxWorkUnits,
        CancellationToken cancellationToken) {
        var workBudget = new PdfUnderstandingWorkBudget(maxWorkUnits, cancellationToken);
        TaggedStructureGraph graph = reads.Graph;
        var selectedPageIndexes = new Dictionary<int, List<int>>(selectedPageNumbers.Length);
        for (int pageIndex = 0; pageIndex < selectedPageNumbers.Length; pageIndex++) {
            workBudget.Consume();
            if (!selectedPageIndexes.TryGetValue(selectedPageNumbers[pageIndex], out List<int>? pageIndexes)) {
                pageIndexes = new List<int>(1);
                selectedPageIndexes.Add(selectedPageNumbers[pageIndex], pageIndexes);
            }
            pageIndexes.Add(pageIndex);
        }
        var contentIndexes = new TaggedPageContentIndex?[selectedPageNumbers.Length];
        var failed = new HashSet<int>();
        var supplied = new HashSet<int>();
        foreach (PdfStructureElementInfo table in graph.Tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(table.ObjectNumber) || !HasResolvedRole(graph.Tagged, table, "Table")) continue;
            if (!reads.Tables.TryGetValue(table.ObjectNumber, out var read)) {
                TaggedTableFailure? failure = TryReadTaggedTable(document, graph, table, selectedPageIndexes, pages, contentIndexes,
                    workBudget, out List<PdfUnderstandingTableCandidate>? candidates, out List<int> touchedPages);
                read = (failure, candidates, touchedPages);
                reads.Tables.Add(table.ObjectNumber, read);
            }
            if (read.Failure is not null) {
                // The same pages its diagnostic will name.
                failed.UnionWith(read.TouchedPages);
                foreach (int pageObject in CollectTablePageObjects(graph, table, workBudget)) {
                    if (document.GetPageNumberForObject(pageObject, cancellationToken) is int page) failed.Add(page);
                }
            } else {
                foreach (PdfUnderstandingTableCandidate candidate in read.Candidates!) supplied.Add(candidate.TaggedSegment!.PageNumber);
            }
        }
        reads.ContentIndexes = (contentIndexes, pages.ToArray());
        var fallBack = new bool[selectedPageNumbers.Length];
        for (int pageIndex = 0; pageIndex < fallBack.Length; pageIndex++) {
            int pageNumber = selectedPageNumbers[pageIndex];
            fallBack[pageIndex] = failed.Contains(pageNumber) || !supplied.Contains(pageNumber);
        }
        return fallBack;
    }

    /// <summary>
    /// Whether every visible text run on the page is bound to an element of the structure tree: it carries a marked-content
    /// id that a reachable element with a role references. Artifacts and whitespace are not content. A page with any other
    /// run holds content the tags don't describe, so its layout can't be taken from them alone.
    /// </summary>
    internal static bool IsFullyTagged(
        PdfReadDocument document,
        TaggedTableReads reads,
        int pageNumber,
        PdfUnderstandingPageResult page,
        long maxWorkUnits,
        CancellationToken cancellationToken) {
        var workBudget = new PdfUnderstandingWorkBudget(maxWorkUnits, cancellationToken);
        if (!reads.RolesBuilt) {
            long consumedBefore = workBudget.Consumed;
            reads.Roles = BuildTaggedContentRoleIndex(document, reads.Graph, workBudget);
            reads.RolesWork = workBudget.Consumed - consumedBefore;
            reads.RolesBuilt = true;
        }
        if (reads.Roles is not TaggedContentRoleIndex roles) return false;
        PdfReadPage readPage = document.Pages[pageNumber - 1];
        bool anyContent = false;
        for (int runIndex = 0; runIndex < page.DecodedRuns.Count; runIndex++) {
            workBudget.Consume();
            PdfTextSpan run = page.DecodedRuns[runIndex];
            if (run.IsArtifactContent || string.IsNullOrWhiteSpace(run.Text)) continue;
            if (run.MarkedContentId is not int mcid ||
                roles.GetRoles(pageNumber, readPage, new MarkedContentKey(run.ContentStreamObjectNumber, mcid)).Count == 0) return false;
            anyContent = true;
        }
        return anyContent;
    }

    /// <summary>
    /// Whether a tagged table segment reads as page layout rather than data: a cell line set noticeably larger than the
    /// table's usual text (a title placed in a cell), a single row below its headers, or cells that are mostly prose.
    /// Producers that lay pages out with tables tag those tables like any other, so the tags alone don't tell them apart.
    /// </summary>
    internal static bool LooksLikeLayout(PdfUnderstandingTableCandidate table) {
        PdfTaggedTableSegment segment = table.TaggedSegment!;
        if (table.Rows.Count - segment.HeaderRowCount <= 1) return true;

        var fontSizes = new List<double>(table.SourceLines.Count);
        for (int index = 0; index < table.SourceLines.Count; index++) {
            double size = table.SourceLines[index].FontSize;
            if (size > 0D && !double.IsInfinity(size)) fontSizes.Add(size);
        }
        if (fontSizes.Count > 0) {
            fontSizes.Sort();
            double median = fontSizes[(fontSizes.Count - 1) / 2];
            if (fontSizes[fontSizes.Count - 1] >= median * HeadingSizedLineRatio) return true;
        }

        int filled = 0;
        int prose = 0;
        foreach (IReadOnlyList<string> row in table.Rows) {
            foreach (string cell in row) {
                if (string.IsNullOrWhiteSpace(cell)) continue;
                filled++;
                if (CountWords(cell) >= ProseCellWords) prose++;
            }
        }
        return filled > 0 && prose * 2 > filled;

        static int CountWords(string text) {
            int words = 0;
            bool inWord = false;
            foreach (char character in text) {
                bool space = char.IsWhiteSpace(character);
                if (!space && !inWord) words++;
                inWord = !space;
            }
            return words;
        }
    }

    /// <summary>A cell line this much larger than the table's median line reads as a heading.</summary>
    private const double HeadingSizedLineRatio = 1.25D;

    /// <summary>A cell with at least this many words reads as prose.</summary>
    private const int ProseCellWords = 8;

    /// <summary>
    /// The page's segments of tables whose tagged read passed, in structure-tree order. The first call indexes every
    /// page in one pass over the structure tree, so asking for each page doesn't walk the tree again.
    /// </summary>
    internal static IReadOnlyList<PdfUnderstandingTableCandidate> GetTaggedTables(TaggedTableReads reads, int pageNumber) {
        if (reads.TablesByPage is null) {
            var byPage = new Dictionary<int, List<PdfUnderstandingTableCandidate>>();
            foreach (PdfStructureElementInfo table in reads.Graph.Tagged.StructureElements) {
                if (!reads.Tables.TryGetValue(table.ObjectNumber, out var read) || read.Failure is not null || read.Candidates is null) continue;
                foreach (PdfUnderstandingTableCandidate candidate in read.Candidates) {
                    int page = candidate.TaggedSegment!.PageNumber;
                    if (!byPage.TryGetValue(page, out var list)) byPage[page] = list = new List<PdfUnderstandingTableCandidate>();
                    list.Add(candidate);
                }
            }
            reads.TablesByPage = byPage;
        }
        return reads.TablesByPage.TryGetValue(pageNumber, out var tables) ? tables : Array.Empty<PdfUnderstandingTableCandidate>();
    }

    /// <summary>
    /// The structure graph and each tagged table's read, kept across the enrichment passes of one document read. A
    /// table's read depends only on the structure tree and the pages' text runs and lines, which a later pass reuses.
    /// </summary>
    internal sealed class TaggedTableReads {
        internal TaggedTableReads(TaggedStructureGraph graph, HashSet<int> pages) {
            Graph = graph;
            Pages = pages;
        }

        internal TaggedStructureGraph Graph { get; }
        /// <summary>Pages a tagged table names.</summary>
        internal HashSet<int> Pages { get; }
        internal Dictionary<int, (TaggedTableFailure? Failure, List<PdfUnderstandingTableCandidate>? Candidates, List<int> TouchedPages)> Tables { get; } = new();
        // Each page's marked-content index from the read, with the analysis it was built from; enrichment reuses it
        // while that analysis stands. The entries are the enricher's own index type.
        internal (object?[] Indexes, PdfUnderstandingPageResult[] Pages)? ContentIndexes { get; set; }
        // Which marked content the structure tree binds to roles, built on first use.
        internal bool RolesBuilt { get; set; }
        internal object? Roles { get; set; }
        // Work units building the role index took, charged again to each later pass that reuses it.
        internal long RolesWork { get; set; }
        // Passing tables' segments by page, built on first use once every table has been read.
        internal Dictionary<int, List<PdfUnderstandingTableCandidate>>? TablesByPage { get; set; }
    }

    private static void ApplyTaggedTableEvidence(
        PdfReadDocument document,
        int[] selectedPageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates,
        TaggedStructureGraph? graph,
        TaggedContentRoleIndex? taggedRoles,
        int maximumArtifactsPerPage,
        List<PdfTaggedTableDiagnostic>?[] diagnostics,
        PdfUnderstandingWorkBudget workBudget,
        TaggedTableReads? reads) {
        // Only a catalog that declares itself tagged (/MarkInfo /Marked true) is trusted for table
        // structure; untagged and unmarked files keep geometry without walking table elements.
        if (graph is null || graph.Tagged.Marked != true) return;

        var selectedPageIndexes = new Dictionary<int, List<int>>(selectedPageNumbers.Length);
        var contentIndexes = new TaggedPageContentIndex?[selectedPageNumbers.Length];
        if (reads?.ContentIndexes is { } built && built.Indexes.Length == contentIndexes.Length) {
            for (int pageIndex = 0; pageIndex < contentIndexes.Length; pageIndex++) {
                if (ReferenceEquals(built.Pages[pageIndex], pages[pageIndex])) contentIndexes[pageIndex] = (TaggedPageContentIndex?)built.Indexes[pageIndex];
            }
        }
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

            TaggedTableFailure? failure;
            List<PdfUnderstandingTableCandidate>? candidates;
            List<int> touchedPages;
            if (reads is not null && reads.Tables.TryGetValue(table.ObjectNumber, out var read)) {
                (failure, candidates, touchedPages) = read;
            } else {
                failure = TryReadTaggedTable(
                    document,
                    graph,
                    table,
                    selectedPageIndexes,
                    pages,
                    contentIndexes,
                    workBudget,
                    out candidates,
                    out touchedPages);
                reads?.Tables.Add(table.ObjectNumber, (failure, candidates, touchedPages));
            }
            if (failure is TaggedTableFailure reported) {
                AddTaggedTableDiagnostic(document, graph, table, reported, touchedPages, selectedPageIndexes, diagnostics, workBudget);
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
                    additions.Add(occurrence == 0 ? candidate : candidate.CopyForOccurrence(workBudget.Consume, workBudget.ThrowIfCancellationRequested));
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
            // Each page's repeated header rows go, in structure order, before that page's first body row.
            var headersByPage = new SortedDictionary<int, List<TaggedGridRow>>();
            foreach (TaggedGridRow repeated in pendingHeaders) {
                workBudget.Consume();
                if (!headersByPage.TryGetValue(repeated.PageNumber, out List<TaggedGridRow>? pageHeaders)) {
                    pageHeaders = new List<TaggedGridRow>();
                    headersByPage.Add(repeated.PageNumber, pageHeaders);
                }
                pageHeaders.Add(repeated);
            }
            foreach (List<TaggedGridRow> pageHeaders in headersByPage.Values) pageHeaders.Sort(static (left, right) => left.GridIndex.CompareTo(right.GridIndex));
            var ordered = new List<TaggedGridRow>(pagedRows.Count + pendingHeaders.Count);
            int lastPage = 0;
            void EmitHeadersThrough(int page) {
                foreach (KeyValuePair<int, List<TaggedGridRow>> pageHeaders in headersByPage) {
                    if (pageHeaders.Key <= lastPage) continue;
                    if (pageHeaders.Key > page) break;
                    workBudget.Consume(pageHeaders.Value.Count);
                    ordered.AddRange(pageHeaders.Value);
                }
            }
            foreach (TaggedGridRow row in pagedRows) {
                workBudget.Consume();
                if (row.PageNumber > lastPage) {
                    EmitHeadersThrough(row.PageNumber);
                    lastPage = row.PageNumber;
                }
                ordered.Add(row);
            }
            EmitHeadersThrough(int.MaxValue);
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
        if (!anySelected) {
            if (grid.All(static row => row.PageNumber == 0)) {
                // No cell references any content. The diagnostic goes to the pages the table names, if selected.
                return new TaggedTableFailure(PdfTaggedTableFallbackReason.NoContent, "The tagged table holds no text.");
            }
            // The table sits on pages outside the selection; it adds nothing to the selected pages.
            candidates = new List<PdfUnderstandingTableCandidate>();
            return null;
        }
        if (!HasConsistentCellOrder(grid, workBudget, out int conflictingRow)) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.CellOrderConflictsWithLayout,
                "Row " + (conflictingRow + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " lists its cells in a different order than they appear on the page.");
        }

        // Drop rows and columns that hold no text anywhere in the selected rows: spreadsheet exports
        // tag unused grid tracks and spacer rows. Empty cells inside a kept row and column remain.
        var keptRows = grid.Where(static row => row.PageNumber != 0 && row.Cells.Any(static cell => cell.Resolved is not null)).ToList();
        ProjectRowSpans(grid, keptRows, workBudget);
        var keptColumn = new bool[columnCount];
        foreach (TaggedGridRow row in keptRows) {
            foreach (TaggedGridCell cell in row.Cells) {
                workBudget.Consume();
                if (cell.Resolved is not null) keptColumn[cell.Column] = true;
            }
        }
        // Columns filled on pages outside the selection belong to the table too, so a selected
        // continuation page keeps them even where its own cells are empty. Content that draws no
        // text counts as empty there, as it does on selected pages.
        var textlessOutsideSelection = new Dictionary<int, Dictionary<int, int?>>();
        foreach (TaggedGridRow row in grid) {
            if (row.PageNumber == 0 || selectedPageIndexes.ContainsKey(row.PageNumber)) continue;
            foreach (TaggedGridCell cell in row.Cells) {
                workBudget.Consume();
                if (!keptColumn[cell.Column] &&
                    HasTextOutsideSelection(document, cell, row.PageNumber, textlessOutsideSelection, workBudget)) {
                    keptColumn[cell.Column] = true;
                }
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
        // The size test covers the whole table: rows on pages outside the selection count by the
        // content their cells reference, so selecting only a short continuation page keeps its grid.
        int rowsWithText = keptRows.Count;
        bool anyRowWithTwoCells = keptRows.Any(static row => row.Cells.Count(static cell => cell.Resolved is not null) >= 2);
        foreach (TaggedGridRow row in grid) {
            workBudget.Consume();
            if (rowsWithText >= 2 && anyRowWithTwoCells) break;
            if (row.PageNumber == 0 || selectedPageIndexes.ContainsKey(row.PageNumber)) continue;
            int cellsWithContent = row.Cells.Count(cell =>
                HasTextOutsideSelection(document, cell, row.PageNumber, textlessOutsideSelection, workBudget));
            if (cellsWithContent == 0) continue;
            rowsWithText++;
            anyRowWithTwoCells |= cellsWithContent >= 2;
        }
        if (keptColumnCount < 2 || rowsWithText < 2 || !anyRowWithTwoCells) {
            return new TaggedTableFailure(PdfTaggedTableFallbackReason.LayoutOnly,
                "The tagged table has fewer than two rows with text or no row with two filled cells, so it is read as page layout.");
        }

        int leadingHeaderRows = LeadingHeaderRows(keptRows, 0);
        IReadOnlyList<string>? header = leadingHeaderRows > 0
            ? CombineHeaderRows(keptRows.Take(leadingHeaderRows).Select(row => BuildTaggedRowTexts(row, columnMap, keptColumnCount)), keptColumnCount)
            : ResolveHeaderOutsideSelection(document, grid, selectedPageIndexes, columnMap, keptColumnCount, workBudget);

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
                    LeadingHeaderRows(keptRows, start),
                    header,
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
    /// Gives each page segment a row span reaches its own empty anchor on the first kept row it
    /// covers there: when the origin row holds no text and was dropped, or sits on another page.
    /// The covered positions then stay one merged cell instead of separate empty cells.
    /// </summary>
    private static void ProjectRowSpans(
        List<TaggedGridRow> grid,
        List<TaggedGridRow> keptRows,
        PdfUnderstandingWorkBudget workBudget) {
        // Pages on which each grid row is kept (header copies keep one grid row on several pages).
        var keptPages = new Dictionary<int, HashSet<int>>();
        foreach (TaggedGridRow row in keptRows) {
            workBudget.Consume();
            if (!keptPages.TryGetValue(row.GridIndex, out HashSet<int>? pages)) keptPages.Add(row.GridIndex, pages = new HashSet<int>());
            pages.Add(row.PageNumber);
        }
        var seen = new HashSet<(int GridIndex, int Column)>();
        var origins = new List<(TaggedGridRow Row, TaggedGridCell Cell)>();
        foreach (TaggedGridRow row in grid) {
            foreach (TaggedGridCell cell in row.Cells) {
                workBudget.Consume();
                if (cell.RowSpan > 1 && seen.Add((row.GridIndex, cell.Column))) origins.Add((row, cell));
            }
        }
        foreach ((TaggedGridRow origin, TaggedGridCell cell) in origins) {
            long end = (long)origin.GridIndex + cell.RowSpan;
            keptPages.TryGetValue(origin.GridIndex, out HashSet<int>? originPages);
            var anchoredPages = new HashSet<int>();
            foreach (TaggedGridRow target in keptRows) {
                workBudget.Consume();
                if (target.GridIndex <= origin.GridIndex || target.GridIndex >= end) continue;
                if (originPages?.Contains(target.PageNumber) == true || !anchoredPages.Add(target.PageNumber)) continue;
                target.Cells.Add(new TaggedGridCell(
                    cell.Element,
                    cell.Column,
                    (int)(end - target.GridIndex),
                    cell.ColumnSpan,
                    cell.IsHeader));
            }
        }
    }

    /// <summary>
    /// Checks that tagged cell order runs the same horizontal direction as the cells on the page in
    /// every row. Either direction is accepted, so right-to-left tables pass; rotated cells are not judged.
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
                // Rotated cells have no horizontal order to judge; they neither set nor break the direction.
                if (current.SourceRuns.Any(static run => Math.Abs(run.RotationDegrees) > 0.5D)) continue;
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

    /// <summary>
    /// Reads the tagged header row's text from its own page when that page is outside the selection,
    /// so selected continuation pages keep their column names.
    /// </summary>
    /// <summary>
    /// Tells whether a cell on a page outside the selection references content that draws text,
    /// using the same textless-content scan that selected pages apply.
    /// </summary>
    private static bool HasTextOutsideSelection(
        PdfReadDocument document,
        TaggedGridCell cell,
        int pageNumber,
        Dictionary<int, Dictionary<int, int?>> textlessByPage,
        PdfUnderstandingWorkBudget workBudget) {
        PdfReadPage? readPage = null;
        Dictionary<int, int?>? textless = null;
        foreach (MarkedContentKey key in cell.ContentOn(pageNumber)) {
            workBudget.Consume();
            readPage ??= document.Pages[pageNumber - 1];
            if (!readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)) return true;
            if (textless is null && !textlessByPage.TryGetValue(pageNumber, out textless)) {
                textless = readPage.GetTextlessPageMarkedContent(() => workBudget.Consume(), workBudget.CancellationToken);
                textlessByPage.Add(pageNumber, textless);
            }
            bool isTextless = textless.TryGetValue(key.MarkedContentId, out int? stream) &&
                (!key.ContentStreamObjectNumber.HasValue || key.ContentStreamObjectNumber == stream);
            if (!isTextless) return true;
        }
        return false;
    }

    private static string[]? ResolveHeaderOutsideSelection(
        PdfReadDocument document,
        List<TaggedGridRow> grid,
        Dictionary<int, List<int>> selectedPageIndexes,
        int[] columnMap,
        int columnCount,
        PdfUnderstandingWorkBudget workBudget) {
        TaggedGridRow? first = grid.FirstOrDefault(static row => row.IsHeader && row.PageNumber != 0);
        if (first is null || selectedPageIndexes.ContainsKey(first.PageNumber)) return null;
        int headerPageNumber = first.PageNumber;
        PdfReadPage headerPage = document.Pages[headerPageNumber - 1];
        IReadOnlyList<PdfTextSpan> spans = headerPage.GetTextSpans(workBudget.CancellationToken);
        var runsByKey = new Dictionary<MarkedContentKey, List<PdfTextSpan>>();
        foreach (PdfTextSpan span in spans) {
            workBudget.Consume();
            if (span.MarkedContentId is not int mcid) continue;
            var key = new MarkedContentKey(span.ContentStreamObjectNumber, mcid);
            if (!runsByKey.TryGetValue(key, out List<PdfTextSpan>? runs)) runsByKey.Add(key, runs = new List<PdfTextSpan>());
            runs.Add(span);
        }
        var headerRows = new List<string[]>();
        foreach (TaggedGridRow header in grid.Where(row => row.IsHeader && row.PageNumber == headerPageNumber)) {
            var texts = new string[columnCount];
            for (int index = 0; index < texts.Length; index++) texts[index] = string.Empty;
            foreach (TaggedGridCell cell in header.Cells) {
                int column = cell.Column < columnMap.Length ? columnMap[cell.Column] : -1;
                if (column < 0) continue;
                var cellRuns = new HashSet<PdfTextSpan>();
                foreach (MarkedContentKey key in cell.ContentOn(header.PageNumber)) {
                    if (TryResolveMarkedRuns(runsByKey, headerPage, key, workBudget, out IReadOnlyList<PdfTextSpan> keyRuns)) {
                        cellRuns.UnionWith(keyRuns);
                    }
                }
                PdfTextSpan[] runs = cellRuns
                    .OrderByDescending(static run => run.Y)
                    .ThenBy(static run => run.X)
                    .ToArray();
                workBudget.Consume(runs.Length + 1L);
                if (runs.Length == 0) continue;
                List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(
                    runs,
                    new TextLayoutEngine.Options { ForceSingleColumn = true },
                    workBudget.Consume,
                    workBudget.ThrowIfCancellationRequested);
                texts[column] = string.Join(" ", lines.Select(static line => line.Text)).Trim();
            }
            headerRows.Add(texts);
        }
        string[] combined = CombineHeaderRows(headerRows, columnCount);
        return combined.Any(static text => text.Length > 0) ? combined : null;
    }

    private static int LeadingHeaderRows(List<TaggedGridRow> rows, int start) {
        int end = start;
        while (end < rows.Count && rows[end].IsHeader && rows[end].PageNumber == rows[start].PageNumber) end++;
        return end - start;
    }

    /// <summary>Joins each column's texts from stacked header rows, top to bottom.</summary>
    private static string[] CombineHeaderRows(IEnumerable<string[]> rows, int columnCount) {
        var parts = new List<string>[columnCount];
        foreach (string[] row in rows) {
            for (int column = 0; column < columnCount && column < row.Length; column++) {
                if (row[column].Length == 0) continue;
                (parts[column] ??= new List<string>()).Add(row[column]);
            }
        }
        var names = new string[columnCount];
        for (int column = 0; column < columnCount; column++) {
            names[column] = parts[column] is null ? string.Empty : string.Join(" ", parts[column]);
        }
        return names;
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
        TaggedStructureGraph graph,
        PdfStructureElementInfo table,
        TaggedTableFailure failure,
        List<int> touchedPages,
        Dictionary<int, List<int>> selectedPageIndexes,
        List<PdfTaggedTableDiagnostic>?[] diagnostics,
        PdfUnderstandingWorkBudget workBudget) {
        // Validation can stop before every row's content is read, and a table spanning pages carries no
        // /Pg of its own, so add the pages named by its descendants and their marked-content references.
        var pagesToReport = new List<int>(touchedPages);
        foreach (int pageObject in CollectTablePageObjects(graph, table, workBudget)) {
            if (document.GetPageNumberForObject(pageObject, workBudget.CancellationToken) is int page &&
                !pagesToReport.Contains(page)) pagesToReport.Add(page);
        }
        pagesToReport.Sort();
        foreach (int pageNumber in pagesToReport) {
            workBudget.Consume();
            if (!selectedPageIndexes.TryGetValue(pageNumber, out List<int>? pageIndexes)) continue;
            foreach (int pageIndex in pageIndexes) {
                (diagnostics[pageIndex] ??= new List<PdfTaggedTableDiagnostic>()).Add(
                    new PdfTaggedTableDiagnostic(table.ObjectNumber, pageNumber, failure.Reason, failure.Message));
            }
        }
    }

    /// <summary>Page objects named by a table and its reachable descendants, tolerating broken parent links.</summary>
    private static List<int> CollectTablePageObjects(
        TaggedStructureGraph graph,
        PdfStructureElementInfo table,
        PdfUnderstandingWorkBudget workBudget) {
        var pages = new List<int>();
        var visited = new HashSet<int>();
        var pending = new Stack<PdfStructureElementInfo>();
        pending.Push(table);
        while (pending.Count > 0) {
            workBudget.Consume();
            PdfStructureElementInfo current = pending.Pop();
            if (!visited.Add(current.ObjectNumber)) continue;
            if (current.PageObjectNumber is int page && !pages.Contains(page)) pages.Add(page);
            foreach (PdfMarkedContentReference reference in current.MarkedContentReferences) {
                workBudget.Consume();
                if (reference.PageObjectNumber is int referencePage && !pages.Contains(referencePage)) pages.Add(referencePage);
            }
            foreach (int child in current.ChildElementObjectNumbers) {
                workBudget.Consume();
                if (graph.StructuresByObject.TryGetValue(child, out PdfStructureElementInfo? childElement)) pending.Push(childElement);
            }
        }
        return pages;
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
        int headerRowCount,
        IReadOnlyList<string>? columnNames,
        PdfUnderstandingWorkBudget workBudget) {
        bool firstRowIsHeader = headerRowCount > 0;
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
                // A cell whose first column was dropped anchors on the first retained column it covers.
                int column = -1;
                for (int spanned = cell.Column; spanned < cell.Column + cell.ColumnSpan && spanned < columnMap.Length; spanned++) {
                    if (columnMap[spanned] >= 0) { column = columnMap[spanned]; break; }
                }
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
            headerRowCount,
            columnNames,
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

    /// <summary>Line options for one cell's text; the layout engine only reads them.</summary>
    private static readonly TextLayoutEngine.Options SingleColumnLines = new() { ForceSingleColumn = true };

    private sealed class TaggedPageContentIndex {
        private readonly PdfReadPage _readPage;
        private readonly PdfUnderstandingPageResult _page;
        private readonly PdfUnderstandingWorkBudget _workBudget;
        private readonly Action<long> _consumeWork;
        private readonly Action _throwIfCancellationRequested;
        private readonly Dictionary<MarkedContentKey, List<PdfTextSpan>> _runsByKey = new();
        private Dictionary<int, int?>? _textlessPageContent;
        // Built on the first cell: the page line holding each run (and any further ones), and the line
        // baselines in ascending order with each one's line index.
        private Dictionary<PdfTextSpan, int>? _lineOwners;
        private Dictionary<PdfTextSpan, List<int>>? _moreLineOwners;
        private double[]? _lineBaselines;
        private int[]? _lineOrder;
        private double _maximumLineTolerance;
        private readonly List<int> _nearLines = new();
        private readonly List<PdfTextSpan> _lineRuns = new();
        // A cell's distinct runs in the order first seen. Cells hold a few runs, compared by reference
        // until there are enough to be worth a set.
        private readonly List<PdfTextSpan> _cellRuns = new();
        private readonly HashSet<PdfTextSpan> _cellRunSet = new();
        private const int CellRunsWithoutSet = 16;

        internal TaggedPageContentIndex(
            PdfReadPage readPage,
            PdfUnderstandingPageResult page,
            PdfUnderstandingWorkBudget workBudget) {
            _readPage = readPage;
            _page = page;
            _workBudget = workBudget;
            _consumeWork = workBudget.Consume;
            _throwIfCancellationRequested = workBudget.ThrowIfCancellationRequested;
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
            _cellRuns.Clear();
            _cellRunSet.Clear();
            for (int keyIndex = 0; keyIndex < keys.Count; keyIndex++) {
                _workBudget.Consume();
                if (!TryResolveRuns(keys[keyIndex], out IReadOnlyList<PdfTextSpan>? runs)) {
                    if (IsTextlessPageContent(keys[keyIndex])) continue;
                    cell = default;
                    return false;
                }
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    _workBudget.Consume();
                    AddCellRun(runs[runIndex]);
                }
            }
            if (_cellRuns.Count == 0) {
                cell = null;
                return true;
            }

            var orderedRuns = new PdfTextSpan[_cellRuns.Count];
            for (int orderedRunIndex = 0; orderedRunIndex < orderedRuns.Length; orderedRunIndex++) {
                _workBudget.Consume();
                orderedRuns[orderedRunIndex] = _cellRuns[orderedRunIndex];
            }
            Array.Sort(orderedRuns, (left, right) => {
                _workBudget.Consume();
                int baseline = right.Y.CompareTo(left.Y);
                return baseline != 0 ? baseline : left.X.CompareTo(right.X);
            });
            List<TextLayoutEngine.TextLine> layoutLines = TextLayoutEngine.BuildLines(
                orderedRuns,
                SingleColumnLines,
                _consumeWork,
                _throwIfCancellationRequested);
            string text = JoinLineTexts(layoutLines);
            if (text.Length == 0) {
                cell = null;
                return orderedRuns.All(static run => string.IsNullOrWhiteSpace(run.Text));
            }

            var sourceLines = new List<PdfUnderstandingLine>();
            List<int> candidateLines = FindLinesNear(orderedRuns);
            for (int candidate = 0; candidate < candidateLines.Count; candidate++) {
                _workBudget.Consume();
                int lineIndex = candidateLines[candidate];
                PdfUnderstandingLine line = _page.Lines[lineIndex];
                double tolerance = LineTolerance(line);
                _lineRuns.Clear();
                for (int runIndex = 0; runIndex < orderedRuns.Length; runIndex++) {
                    _workBudget.Consume();
                    PdfTextSpan run = orderedRuns[runIndex];
                    if (RunSitsOnLine(run, lineIndex, line, tolerance)) _lineRuns.Add(run);
                }
                if (_lineRuns.Count == 0) continue;
                PdfTextSpan[] lineRuns;
                string lineText;
                if (_lineRuns.Count == orderedRuns.Length) {
                    // Every run of the cell sits on this line, in the same order: its text is the cell's text.
                    lineRuns = orderedRuns;
                    lineText = text;
                } else {
                    lineRuns = _lineRuns.ToArray();
                    lineText = JoinLineTexts(TextLayoutEngine.BuildLines(
                        lineRuns,
                        SingleColumnLines,
                        _consumeWork,
                        _throwIfCancellationRequested));
                    if (lineText.Length == 0) continue;
                }
                PdfDoubleAggregate left = default, right = default, baseline = default, fontSize = default, rotation = default, advance = default;
                for (int runIndex = 0; runIndex < lineRuns.Length; runIndex++) {
                    PdfTextSpan run = lineRuns[runIndex];
                    left.Add(run.X);
                    right.Add(run.X + Math.Max(0D, run.Advance));
                    baseline.Add(run.Y);
                    fontSize.Add(run.FontSize);
                    rotation.Add(run.RotationDegrees);
                    advance.Add(Math.Max(0D, run.Advance));
                }
                var word = new PdfUnderstandingWord(
                    lineText,
                    left.Minimum,
                    right.Maximum,
                    baseline.Average,
                    fontSize.Maximum,
                    rotation.Average,
                    Array.AsReadOnly(lineRuns),
                    line.Confidence,
                    line.Evidence,
                    advance.Sum,
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

            PdfDoubleAggregate cellLeft = default, cellRight = default, cellY = default;
            for (int runIndex = 0; runIndex < orderedRuns.Length; runIndex++) {
                PdfTextSpan run = orderedRuns[runIndex];
                cellLeft.Add(run.X);
                cellRight.Add(run.X + Math.Max(0D, run.Advance));
                cellY.Add(run.Y);
            }
            cell = new TaggedTableCell(
                text,
                cellLeft.Minimum,
                cellRight.Maximum,
                cellY.Maximum,
                cellY.Minimum,
                sourceLines.AsReadOnly(),
                Array.AsReadOnly(orderedRuns));
            return true;
        }

        private void AddCellRun(PdfTextSpan run) {
            if (_cellRuns.Count < CellRunsWithoutSet) {
                for (int index = 0; index < _cellRuns.Count; index++) {
                    if (ReferenceEquals(_cellRuns[index], run)) return;
                }
                _cellRuns.Add(run);
                if (_cellRuns.Count == CellRunsWithoutSet) _cellRunSet.UnionWith(_cellRuns);
                return;
            }
            if (_cellRunSet.Add(run)) _cellRuns.Add(run);
        }

        /// <summary>The lines' texts joined by spaces and trimmed, as one cell value.</summary>
        private static string JoinLineTexts(List<TextLayoutEngine.TextLine> lines) {
            if (lines.Count == 1) return lines[0].Text.Trim();
            var builder = new System.Text.StringBuilder();
            for (int index = 0; index < lines.Count; index++) {
                if (index > 0) builder.Append(' ');
                builder.Append(lines[index].Text);
            }
            return builder.ToString().Trim();
        }

        private static double LineTolerance(PdfUnderstandingLine line) => Math.Max(2.5D, line.FontSize * 0.35D);

        /// <summary>
        /// Indexes, in page order, of the page lines a cell's runs sit on: the lines that hold each run, or for a run
        /// no line holds, every line whose baseline may lie within its tolerance of the run.
        /// </summary>
        private List<int> FindLinesNear(PdfTextSpan[] runs) {
            EnsureLineIndexes();
            _nearLines.Clear();
            for (int runIndex = 0; runIndex < runs.Length; runIndex++) {
                _workBudget.Consume();
                PdfTextSpan run = runs[runIndex];
                if (_lineOwners!.TryGetValue(run, out int owner)) {
                    _nearLines.Add(owner);
                    if (_moreLineOwners is not null && _moreLineOwners.TryGetValue(run, out List<int>? others)) _nearLines.AddRange(others);
                    continue;
                }
                AddLinesNearBaseline(run.Y);
            }
            if (_nearLines.Count > 1) {
                _nearLines.Sort();
                int kept = 1;
                for (int index = 1; index < _nearLines.Count; index++) {
                    if (_nearLines[index] != _nearLines[kept - 1]) _nearLines[kept++] = _nearLines[index];
                }
                _nearLines.RemoveRange(kept, _nearLines.Count - kept);
            }
            return _nearLines;
        }

        /// <summary>A run held by page lines sits on those lines only; any other run sits on every line near its baseline.</summary>
        private bool RunSitsOnLine(PdfTextSpan run, int lineIndex, PdfUnderstandingLine line, double tolerance) {
            if (!_lineOwners!.TryGetValue(run, out int owner)) return Math.Abs(run.Y - line.BaselineY) <= tolerance;
            return owner == lineIndex ||
                _moreLineOwners is not null && _moreLineOwners.TryGetValue(run, out List<int>? others) && others.Contains(lineIndex);
        }

        private void AddLinesNearBaseline(double y) {
            if (double.IsNaN(y)) return;
            double from = y - _maximumLineTolerance;
            double to = y + _maximumLineTolerance;
            if (double.IsNaN(from)) from = double.NegativeInfinity;
            if (double.IsNaN(to)) to = double.PositiveInfinity;
            double[] keys = _lineBaselines!;
            int low = 0, high = keys.Length;
            while (low < high) {
                int middle = low + ((high - low) >> 1);
                if (keys[middle] < from) low = middle + 1;
                else high = middle;
            }
            for (int index = low; index < keys.Length && keys[index] <= to; index++) {
                _workBudget.Consume();
                _nearLines.Add(_lineOrder![index]);
            }
        }

        private void EnsureLineIndexes() {
            if (_lineOwners is not null) return;
            var baselines = new List<double>(_page.Lines.Count);
            var order = new List<int>(_page.Lines.Count);
            var owners = new Dictionary<PdfTextSpan, int>(_page.DecodedRuns.Count);
            double maximumTolerance = 0D;
            for (int lineIndex = 0; lineIndex < _page.Lines.Count; lineIndex++) {
                _workBudget.Consume();
                PdfUnderstandingLine line = _page.Lines[lineIndex];
                for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                    IReadOnlyList<PdfTextSpan> wordRuns = line.Words[wordIndex].SourceRuns;
                    for (int runIndex = 0; runIndex < wordRuns.Count; runIndex++) {
                        _workBudget.Consume();
                        PdfTextSpan run = wordRuns[runIndex];
                        if (!owners.TryGetValue(run, out int owner)) {
                            owners.Add(run, lineIndex);
                        } else if (owner != lineIndex) {
                            _moreLineOwners ??= new Dictionary<PdfTextSpan, List<int>>();
                            if (!_moreLineOwners.TryGetValue(run, out List<int>? others)) _moreLineOwners.Add(run, others = new List<int>(1));
                            if (!others.Contains(lineIndex)) others.Add(lineIndex);
                        }
                    }
                }
                double tolerance = LineTolerance(line);
                if (double.IsNaN(tolerance) || double.IsNaN(line.BaselineY)) continue;
                baselines.Add(line.BaselineY);
                order.Add(lineIndex);
                maximumTolerance = Math.Max(maximumTolerance, tolerance);
            }
            _lineBaselines = baselines.ToArray();
            _lineOrder = order.ToArray();
            Array.Sort(_lineBaselines, _lineOrder);
            _maximumLineTolerance = maximumTolerance;
            _lineOwners = owners;
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
            out IReadOnlyList<PdfTextSpan> runs) =>
            TryResolveMarkedRuns(_runsByKey, _readPage, key, _workBudget, out runs);
    }

    /// <summary>
    /// Finds the runs for a marked-content key: the exact stream first, otherwise the one compatible
    /// page-stream match, so the same MCID in a Form XObject never lends its text to page content.
    /// </summary>
    private static bool TryResolveMarkedRuns(
        Dictionary<MarkedContentKey, List<PdfTextSpan>> runsByKey,
        PdfReadPage readPage,
        MarkedContentKey key,
        PdfUnderstandingWorkBudget workBudget,
        out IReadOnlyList<PdfTextSpan> runs) {
        if (runsByKey.TryGetValue(key, out List<PdfTextSpan>? exact)) {
            runs = exact;
            return exact.Count > 0;
        }

        List<PdfTextSpan>? match = null;
        foreach (KeyValuePair<MarkedContentKey, List<PdfTextSpan>> candidate in runsByKey) {
            workBudget.Consume();
            if (candidate.Key.MarkedContentId != key.MarkedContentId) continue;
            bool compatible = key.ContentStreamObjectNumber.HasValue
                ? candidate.Key.ContentStreamObjectNumber is null &&
                  readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)
                : readPage.IsPageContentStreamObjectNumber(candidate.Key.ContentStreamObjectNumber);
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

    private readonly record struct TaggedTableRowSource(
        bool IsHeader,
        List<PdfStructureElementInfo> Cells,
        List<bool> HeaderCells);

    internal readonly record struct TaggedTableFailure(PdfTaggedTableFallbackReason Reason, string Message);

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
