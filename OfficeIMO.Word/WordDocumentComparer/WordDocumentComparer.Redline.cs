using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        /// <summary>Writes a DOCX redline for two paths using the default comparison and redline settings.</summary>
        public static WordComparisonResult CreateRedlineDocument(string sourcePath, string targetPath, string outputPath) =>
            CreateRedlineDocument(sourcePath, targetPath, outputPath, null);

        /// <summary>Compares two documents and writes a DOCX redline document using the configured redline mode.</summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <param name="outputPath">Path where the redline DOCX should be written.</param>
        /// <param name="options">Optional redline and comparison settings.</param>
        /// <returns>The structured comparison result used to generate the redline document.</returns>
        public static WordComparisonResult CreateRedlineDocument(string sourcePath, string targetPath, string outputPath, WordComparisonRedlineOptions? options) {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));

            ValidateRedlineOutputPath(sourcePath, targetPath, outputPath);
            using WordDocument source = WordDocument.Load(sourcePath, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            using WordDocument target = WordDocument.Load(targetPath, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            return CreateRedlineDocumentCore(source, target, outputPath, options, sourcePath, targetPath);
        }

        /// <summary>Writes a redline from the current loaded models, including unsaved edits, without saving, reopening or disposing either input.</summary>
        /// <param name="source">Caller-owned original document, already authenticated if encrypted.</param>
        /// <param name="target">Caller-owned modified document, already authenticated if encrypted.</param>
        /// <param name="outputPath">Destination distinct from either model's associated file.</param>
        /// <param name="options">Native comparison and redline settings. Null selects defaults; supplied settings are not modified.</param>
        /// <returns>The structured comparison used to generate the redline.</returns>
        /// <remarks>InPlaceTarget clones the live target package in memory. Only the requested output is written; authenticated inputs are not staged as plaintext files.</remarks>
        public static WordComparisonResult CreateRedlineDocument(WordDocument source, WordDocument target, string outputPath, WordComparisonRedlineOptions? options) {
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(target);
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));
            ValidateRedlineOutputPath(source.FilePath, target.FilePath, outputPath);
            return CreateRedlineDocumentCore(source, target, outputPath, options, source.FilePath ?? "source", target.FilePath ?? "target");
        }

        private static WordComparisonResult CreateRedlineDocumentCore(WordDocument source, WordDocument target, string outputPath,
            WordComparisonRedlineOptions? options, string sourceLabel, string targetLabel) {
            options ??= new WordComparisonRedlineOptions();
            WordComparisonResult result = CompareStructure(source, target, options.ComparisonOptions, sourceLabel, targetLabel);
            if (options.Mode == WordComparisonRedlineMode.InPlaceTarget &&
                result.Findings.Any(finding => finding.Scope == WordComparisonScope.Shape)) {
                result.AddLimitation(new WordComparisonLimitation(
                    "redline-shape-in-place-text-fallback",
                    "DrawingML shape findings are represented by tracked textual evidence in in-place redlines; native drawing move or replacement revision markup is not synthesized.",
                    result.Findings.Any(finding => finding.Scope == WordComparisonScope.Shape && finding.ChangeKind != WordComparisonChangeKind.Inserted),
                    result.Findings.Any(finding => finding.Scope == WordComparisonScope.Shape && finding.ChangeKind != WordComparisonChangeKind.Deleted)));
            }

            EnsureOutputDirectory(outputPath);

            if (options.Mode == WordComparisonRedlineMode.InPlaceTarget) {
                CreateInPlaceTargetRedlineDocument(source, target, outputPath, result, options);
                return result;
            }

            using WordDocument document = WordDocument.Create(outputPath);
            document.AddParagraph("Word Comparison Redline").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph(WordComparisonReportWriter.ToTextSummary(result));

            if (options.IncludeSummary) {
                AppendRedlineSummary(document, result);
            }

            if (options.IncludeFindingsTable) {
                AppendRedlineFindingsTable(document, result);
            }

            if (ShouldAppendTrackedRedlineFindings(options)) {
                AppendTrackedRedlineFindings(document, result, options);
            }

            document.Save();
            return result;
        }

        private static void ValidateRedlineOutputPath(string? sourcePath, string? targetPath, string outputPath) {
            string outputFullPath = Path.GetFullPath(outputPath);

            if (!string.IsNullOrEmpty(sourcePath) && OfficeIMO.Internal.OfficePathIdentity.AreEquivalent(outputFullPath, sourcePath!)) {
                throw new InvalidOperationException("Redline output must be written to a different path than the source document.");
            }

            if (!string.IsNullOrEmpty(targetPath) && OfficeIMO.Internal.OfficePathIdentity.AreEquivalent(outputFullPath, targetPath!)) {
                throw new InvalidOperationException("Redline output must be written to a different path than the target document.");
            }
        }

        private static void EnsureOutputDirectory(string outputPath) {
            string? directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }
        }

        private static void CreateInPlaceTargetRedlineDocument(WordDocument sourceDocument, WordDocument targetDocument, string outputPath, WordComparisonResult result, WordComparisonRedlineOptions options) {
            using WordDocument document = WordPackageView.CloneForRedline(targetDocument);
            // Source elements are copied into the SDK output with the table grid state an SDK load gives them.
            sourceDocument.OpenSdkPackage("Document comparison: in-place redline");
            WordPackageView source = WordPackageView.For(sourceDocument);
            WordPackageView target = WordPackageView.For(document);
            HashSet<int> rewrittenParagraphs = ApplyParagraphFindings(
                source,
                target,
                result,
                options,
                out IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
                out IReadOnlyList<RedlineParagraphEntry> targetParagraphs);
            ApplyRunFindings(sourceParagraphs, targetParagraphs, result, options, rewrittenParagraphs);
            ApplyContentControlFindings(source, target, result, options);
            ApplyImageFindings(source, target, result, options);

            ApplyTableCellFindings(source, target, result, options);
            ApplyTableRowFindings(source, target, result, options);
            ApplyTableFindings(source, target, result, options);
            AppendInPlaceFeatureAndReviewFindings(document, result, options);

            document.Save(outputPath,
                new WordSaveOptions { SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup });
        }

        private static bool HasTrackedText(WordComparisonFinding finding) {
            return !string.IsNullOrEmpty(finding.SourceText) || !string.IsNullOrEmpty(finding.TargetText);
        }

        private static void RewriteParagraphWithTrackedText(OfficeXmlNode paragraph, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            foreach (OfficeXmlNode child in paragraph.Children().ToList().Where(child => !child.Is(W, "pPr") && !ShouldPreserveParagraphRedlineChild(child)).ToList()) {
                child.Remove();
            }

            OfficeXmlNode insertionPoint = paragraph.Children().ToList().FirstOrDefault(child => !child.Is(W, "pPr"));
            if (!string.IsNullOrEmpty(sourceText)) {
                paragraph.InsertBefore(CreateDeletedRun(paragraph, sourceText!, options), insertionPoint);
            }

            if (!string.IsNullOrEmpty(targetText)) {
                paragraph.InsertBefore(CreateInsertedRun(paragraph, targetText!, options), insertionPoint);
            }
        }

        private static bool ShouldPreserveParagraphRedlineChild(OfficeXmlNode child) {
            if (HasDescendant(child, W, "t") || HasDescendant(child, W, "delText")) {
                return false;
            }

            if (!child.Is(W, "r")) {
                return true;
            }

            return child.Children().ToList().Any(runChild => !runChild.Is(W, "rPr") && !runChild.Is(W, "t"));
        }

        private static void ApplyRunFinding(OfficeXmlNode paragraph, int runIndex, WordComparisonFinding finding, WordComparisonRedlineOptions options) {
            List<OfficeXmlNode> runs = GetDirectParagraphRuns(paragraph);
            if (runs.Count == 0) {
                RewriteParagraphWithTrackedText(paragraph, finding.SourceText, finding.TargetText, options);
                return;
            }

            int boundedRunIndex = Math.Max(0, Math.Min(runIndex, runs.Count - 1));
            OfficeXmlNode targetRun = runs[boundedRunIndex];
            OfficeXmlNode parent = targetRun.Parent.IsNone ? paragraph : targetRun.Parent;
            bool insertTrailingDeletion = finding.ChangeKind == WordComparisonChangeKind.Deleted && runIndex >= runs.Count;

            if (finding.ChangeKind != WordComparisonChangeKind.Inserted && !string.IsNullOrEmpty(finding.SourceText)) {
                OfficeXmlNode deletedRun = CreateDeletedRun(parent, finding.SourceText!, options);
                if (insertTrailingDeletion) {
                    targetRun.InsertAfterSelf(deletedRun);
                } else {
                    parent.InsertBefore(deletedRun, targetRun);
                }
            }

            if (finding.ChangeKind != WordComparisonChangeKind.Deleted && !string.IsNullOrEmpty(finding.TargetText)) {
                parent.InsertBefore(CreateInsertedRun(parent, finding.TargetText!, options), targetRun);
            }

            if (finding.ChangeKind != WordComparisonChangeKind.Deleted) {
                targetRun.Remove();
            }
        }

        private static void ApplyTableCellFindings(WordPackageView sourceDocument, WordPackageView targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            HashSet<string> nestedTableParentCellKeys = GetNestedTableParentCellKeys(sourceTables, targetTables, result, options);
            var rewrittenCells = new HashSet<string>(StringComparer.Ordinal);
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.TableCell ||
                    !TryParseTableCellLocation(finding.Location, out int tableIndex, out int rowIndex, out int cellIndex) ||
                    tableIndex < 0 ||
                    tableIndex >= targetTables.Count ||
                    rowIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Modified:
                    case WordComparisonChangeKind.Inserted:
                        OfficeXmlNode table = targetTables[tableIndex].Table;
                        List<OfficeXmlNode> rows = ChildrenW(table, "tr");
                        if (rowIndex >= rows.Count) {
                            continue;
                        }

                        OfficeXmlNode row = rows[rowIndex];
                        List<OfficeXmlNode> cells = ChildrenW(row, "tc");
                        string cellKey = CreateTableCellKey(tableIndex, rowIndex, cellIndex);
                        if (rewrittenCells.Contains(cellKey) || nestedTableParentCellKeys.Contains(cellKey)) {
                            continue;
                        }

                        if (cellIndex >= 0 && cellIndex < cells.Count) {
                            RewriteCellWithTrackedText(cells[cellIndex], finding.SourceText, finding.TargetText, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenCells.Add(cellKey);
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (!string.IsNullOrEmpty(finding.SourceText)) {
                            RedlineTableEntry targetTable = targetTables[tableIndex];
                            RedlineTableEntry? sourceTable = FindSourceTableForDeletedCell(sourceTables, targetTable, rowIndex, cellIndex, finding.SourceText);
                            if (sourceTable == null) {
                                break;
                            }

                            List<OfficeXmlNode> sourceRows = ChildrenW(sourceTable.Table, "tr");
                            if (rowIndex < 0 || rowIndex >= sourceRows.Count) {
                                break;
                            }

                            List<OfficeXmlNode> targetRows = ChildrenW(targetTable.Table, "tr");
                            int targetRowIndex = FindTargetRowIndexForDeletedCell(sourceRows[rowIndex], targetRows, rowIndex);
                            if (targetRowIndex < 0 || targetRowIndex >= targetRows.Count) {
                                break;
                            }

                            OfficeXmlNode targetRow = targetRows[targetRowIndex];
                            List<OfficeXmlNode> targetCells = ChildrenW(targetRow, "tc");
                            int targetCellIndex = cellIndex;
                            List<OfficeXmlNode> sourceCells = ChildrenW(sourceRows[rowIndex], "tc");
                            if (cellIndex >= 0 && cellIndex < sourceCells.Count) {
                                targetCellIndex = FindTargetGapByNeighborIdentity(sourceCells, targetCells, cellIndex, GetOpenXmlCellText, GetOpenXmlCellText);
                            }

                            string deletedCellKey = CreateTableCellKey(tableIndex, targetRowIndex, targetCellIndex) + "/deleted/" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            if (rewrittenCells.Contains(deletedCellKey) || nestedTableParentCellKeys.Contains(CreateTableCellKey(tableIndex, targetRowIndex, targetCellIndex))) {
                                break;
                            }

                            OfficeXmlNode sourceCell = cellIndex >= 0 && cellIndex < sourceCells.Count ? sourceCells[cellIndex] : default;
                            InsertDeletedCell(targetRow, targetCells, targetCellIndex, sourceCell, finding.SourceText!, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenCells.Add(deletedCellKey);
                        }

                        break;
                }
            }
        }

        private static RedlineTableEntry? FindSourceTableForDeletedCell(
            IReadOnlyList<RedlineTableEntry> sourceTables,
            RedlineTableEntry targetTable,
            int rowIndex,
            int cellIndex,
            string? sourceCellText) {
            RedlineTableEntry? ordinalMatch = sourceTables.FirstOrDefault(table =>
                string.Equals(table.PartKey, targetTable.PartKey, StringComparison.Ordinal) &&
                table.LocalIndex == targetTable.LocalIndex);

            RedlineTableEntry? textMatch = sourceTables
                .Where(table => string.Equals(table.PartKey, targetTable.PartKey, StringComparison.Ordinal))
                .Where(table => {
                    List<OfficeXmlNode> rows = ChildrenW(table.Table, "tr");
                    if (rowIndex < 0 || rowIndex >= rows.Count) {
                        return false;
                    }

                    List<OfficeXmlNode> cells = ChildrenW(rows[rowIndex], "tc");
                    return cellIndex >= 0 &&
                        cellIndex < cells.Count &&
                        string.Equals(GetOpenXmlCellText(cells[cellIndex]), sourceCellText ?? string.Empty, StringComparison.Ordinal);
                })
                .OrderByDescending(table => GetRedlineTableSimilarity(table.Table, targetTable.Table))
                .FirstOrDefault();

            return textMatch ?? ordinalMatch;
        }

        private static int FindTargetRowIndexForDeletedCell(OfficeXmlNode sourceRow, IReadOnlyList<OfficeXmlNode> targetRows, int sourceRowIndex) {
            if (targetRows.Count == 0) {
                return -1;
            }

            string sourceText = GetOpenXmlRowText(sourceRow);
            int bestIndex = -1;
            double bestSimilarity = double.MinValue;
            for (int index = 0; index < targetRows.Count; index++) {
                double similarity = GetTextSimilarity(sourceText, GetOpenXmlRowText(targetRows[index]));
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestIndex >= 0
                ? bestIndex
                : Math.Max(0, Math.Min(sourceRowIndex, targetRows.Count - 1));
        }

        private static HashSet<string> GetNestedTableParentCellKeys(IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, WordComparisonResult result, WordComparisonRedlineOptions options) {
            var cellKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Table ||
                    !TryParseTableLocation(finding.Location, out int tableIndex)) {
                    continue;
                }

                IReadOnlyList<RedlineTableEntry> tables = finding.ChangeKind == WordComparisonChangeKind.Deleted ? sourceTables : targetTables;
                if (tableIndex < 0 || tableIndex >= tables.Count) {
                    continue;
                }

                if (TryGetNestedTablePlacement(tables, tables[tableIndex].Table, out NestedTablePlacement placement)) {
                    cellKeys.Add(CreateTableCellKey(placement.ParentTableIndex, placement.RowIndex, placement.CellIndex));
                }
            }

            return cellKeys;
        }

        private static void ApplyTableRowFindings(WordPackageView sourceDocument, WordPackageView targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            var rewrittenRows = new HashSet<string>(StringComparer.Ordinal);

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.TableRow ||
                    !TryParseTableRowLocation(finding.Location, out int tableIndex, out int rowIndex) ||
                    tableIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                string rowKey = tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                                rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (rewrittenRows.Contains(rowKey)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Inserted:
                        if (tableIndex < targetTables.Count) {
                            List<OfficeXmlNode> targetRows = ChildrenW(targetTables[tableIndex].Table, "tr");
                            if (rowIndex >= 0 && rowIndex < targetRows.Count) {
                                RewriteRowWithTrackedText(targetRows[rowIndex], trackInserted: true, options);
                                rewrittenRows.Add(rowKey);
                            }
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (tableIndex < targetTables.Count) {
                            RedlineTableEntry? sourceTable = FindSourceTableForDeletedRow(sourceTables, targetTables[tableIndex], rowIndex, finding.SourceText);
                            if (sourceTable == null) {
                                break;
                            }

                            List<OfficeXmlNode> sourceRows = ChildrenW(sourceTable.Table, "tr");
                            if (rowIndex >= 0 && rowIndex < sourceRows.Count) {
                                List<OfficeXmlNode> targetRows = ChildrenW(targetTables[tableIndex].Table, "tr");
                                int targetRowIndex = FindTargetGapByNeighborIdentity(sourceRows, targetRows, rowIndex, GetOpenXmlRowText, GetOpenXmlRowText);
                                InsertDeletedRow(targetTables[tableIndex].Table, sourceRows[rowIndex], targetRowIndex, options);
                                rewrittenRows.Add(rowKey);
                            }
                        }

                        break;
                }
            }
        }

        private static RedlineTableEntry? FindSourceTableForDeletedRow(
            IReadOnlyList<RedlineTableEntry> sourceTables,
            RedlineTableEntry targetTable,
            int rowIndex,
            string? sourceRowText) {
            RedlineTableEntry? ordinalMatch = sourceTables.FirstOrDefault(table =>
                string.Equals(table.PartKey, targetTable.PartKey, StringComparison.Ordinal) &&
                table.LocalIndex == targetTable.LocalIndex);

            RedlineTableEntry? textMatch = sourceTables
                .Where(table => string.Equals(table.PartKey, targetTable.PartKey, StringComparison.Ordinal))
                .Where(table => {
                    List<OfficeXmlNode> rows = ChildrenW(table.Table, "tr");
                    return rowIndex >= 0 &&
                        rowIndex < rows.Count &&
                        string.Equals(GetOpenXmlRowText(rows[rowIndex]), sourceRowText ?? string.Empty, StringComparison.Ordinal);
                })
                .OrderByDescending(table => GetRedlineTableSimilarity(table.Table, targetTable.Table))
                .FirstOrDefault();

            return textMatch ?? ordinalMatch;
        }

        private static double GetRedlineTableSimilarity(OfficeXmlNode sourceTable, OfficeXmlNode targetTable) {
            return GetTextSimilarity(GetOpenXmlTableText(sourceTable), GetOpenXmlTableText(targetTable));
        }

        private static int FindTargetGapByNeighborIdentity<TSource, TTarget>(
            IReadOnlyList<TSource> sourceItems,
            IReadOnlyList<TTarget> targetItems,
            int sourceIndex,
            Func<TSource, string> sourceIdentity,
            Func<TTarget, string> targetIdentity) {
            if (sourceIndex < 0 || sourceIndex >= sourceItems.Count || targetItems.Count == 0) {
                return Math.Max(0, Math.Min(sourceIndex, targetItems.Count));
            }

            for (int index = sourceIndex + 1; index < sourceItems.Count; index++) {
                string identity = sourceIdentity(sourceItems[index]);
                if (string.IsNullOrEmpty(identity)) {
                    continue;
                }

                int targetIndex = FindFirstTargetIdentity(targetItems, targetIdentity, identity);
                if (targetIndex >= 0) {
                    return targetIndex;
                }
            }

            for (int index = sourceIndex - 1; index >= 0; index--) {
                string identity = sourceIdentity(sourceItems[index]);
                if (string.IsNullOrEmpty(identity)) {
                    continue;
                }

                int targetIndex = FindLastTargetIdentity(targetItems, targetIdentity, identity);
                if (targetIndex >= 0) {
                    return targetIndex + 1;
                }
            }

            return Math.Max(0, Math.Min(sourceIndex, targetItems.Count));
        }

        private static int FindFirstTargetIdentity<TTarget>(IReadOnlyList<TTarget> targetItems, Func<TTarget, string> targetIdentity, string identity) {
            for (int index = 0; index < targetItems.Count; index++) {
                if (string.Equals(targetIdentity(targetItems[index]), identity, StringComparison.Ordinal)) {
                    return index;
                }
            }

            return -1;
        }

        private static int FindLastTargetIdentity<TTarget>(IReadOnlyList<TTarget> targetItems, Func<TTarget, string> targetIdentity, string identity) {
            for (int index = targetItems.Count - 1; index >= 0; index--) {
                if (string.Equals(targetIdentity(targetItems[index]), identity, StringComparison.Ordinal)) {
                    return index;
                }
            }

            return -1;
        }

        private static string GetOpenXmlTableText(OfficeXmlNode table) {
            return string.Join(TableRowSeparator, ChildrenW(table, "tr").Select(GetOpenXmlRowText).ToArray());
        }

        private static string GetOpenXmlRowText(OfficeXmlNode row) {
            return string.Join(" | ", ChildrenW(row, "tc").Select(GetOpenXmlCellText).ToArray());
        }

        private static List<RedlineTableEntry> GetRedlineTableEntries(WordPackageView document) {
            var entries = new List<RedlineTableEntry>();
            WordPartView? mainPart = document.Main;
            AddRedlineTableEntries(entries, BodyPartKey, BodyOf(mainPart), BodyPartOrderBase);

            if (mainPart == null) {
                return entries;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                AddRedlineTableEntries(entries, headerPartKey.Value, HeaderRoot(headerPartKey.Key), HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                AddRedlineTableEntries(entries, footerPartKey.Value, FooterRoot(footerPartKey.Key), FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                footerIndex++;
            }

            List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                AddRedlineTableEntries(entries, FootnotePartKeyPrefix + noteId, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
            }

            List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                AddRedlineTableEntries(entries, EndnotePartKeyPrefix + noteId, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
            }

            return entries;
        }

        private static void AddRedlineTableEntries(List<RedlineTableEntry> entries, string partKey, OfficeXmlNode container, int orderBase) {
            if (container.IsNone) {
                return;
            }

            int localIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element.Is(W, "tbl")) {
                    entries.Add(new RedlineTableEntry(partKey, container, ordered.Element, localIndex));
                    localIndex++;
                }
            }
        }

        private static void RewriteCellWithTrackedText(OfficeXmlNode cell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            OfficeXmlNode properties = cell.Child(W, "tcPr");
            properties = properties.IsNone ? default : properties.Clone();
            RemoveAllChildren(cell);
            if (!properties.IsNone) {
                cell.Append(properties);
            }

            AppendCellTrackedParagraphs(cell, sourceText, targetText, options);
            EnsureCellHasParagraph(cell);
            RemoveEmptyWordColorAttributes(cell);
        }

        private static void InsertDeletedCell(OfficeXmlNode row, IReadOnlyList<OfficeXmlNode> existingCells, int cellIndex, OfficeXmlNode sourceCell, string sourceText, WordComparisonRedlineOptions options) {
            OfficeXmlNode deletedCell = row.CreateElement(W, "tc");
            OfficeXmlNode properties = sourceCell.IsNone ? default : sourceCell.Child(W, "tcPr");
            if (!properties.IsNone) {
                deletedCell.Append(properties.Clone());
            }

            AppendCellTrackedParagraphs(deletedCell, sourceText, null, options);
            EnsureCellHasParagraph(deletedCell);

            if (cellIndex >= 0 && cellIndex < existingCells.Count) {
                existingCells[cellIndex].Parent.InsertBefore(deletedCell, existingCells[cellIndex]);
            } else {
                row.Append(deletedCell);
            }
        }

        private static void RewriteRowWithTrackedText(OfficeXmlNode row, bool trackInserted, WordComparisonRedlineOptions options) {
            foreach (OfficeXmlNode cell in ChildrenW(row, "tc")) {
                string cellText = GetOpenXmlCellText(cell);
                if (trackInserted) {
                    RewriteCellWithTrackedText(cell, null, cellText, options);
                } else {
                    RewriteCellWithTrackedText(cell, cellText, null, options);
                }
            }
        }

        private static void InsertDeletedRow(OfficeXmlNode targetTable, OfficeXmlNode sourceRow, int rowIndex, WordComparisonRedlineOptions options) {
            OfficeXmlNode deletedRow = sourceRow.Clone();
            WordPackageView.CopyGridNormalizationState(sourceRow, deletedRow);
            RewriteRowWithTrackedText(deletedRow, trackInserted: false, options);

            List<OfficeXmlNode> targetRows = ChildrenW(targetTable, "tr");
            if (rowIndex >= 0 && rowIndex < targetRows.Count) {
                targetRows[rowIndex].Parent.InsertBefore(deletedRow, targetRows[rowIndex]);
            } else {
                targetTable.Append(deletedRow);
            }
        }

        private static void AppendCellTrackedParagraphs(OfficeXmlNode cell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            string[] sourceParagraphs = SplitCellParagraphText(sourceText);
            string[] targetParagraphs = SplitCellParagraphText(targetText);
            int paragraphCount = Math.Max(sourceParagraphs.Length, targetParagraphs.Length);
            for (int index = 0; index < paragraphCount; index++) {
                string? sourceParagraph = index < sourceParagraphs.Length ? sourceParagraphs[index] : null;
                string? targetParagraph = index < targetParagraphs.Length ? targetParagraphs[index] : null;
                OfficeXmlNode paragraph = cell.CreateElement(W, "p");
                if (!string.IsNullOrEmpty(sourceParagraph)) {
                    paragraph.Append(CreateDeletedRun(paragraph, sourceParagraph!, options));
                }

                if (!string.IsNullOrEmpty(targetParagraph)) {
                    paragraph.Append(CreateInsertedRun(paragraph, targetParagraph!, options));
                }

                cell.Append(paragraph);
            }
        }

        private static string[] SplitCellParagraphText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return Array.Empty<string>();
            }

            return text!.Split(new[] { CellParagraphSeparator }, StringSplitOptions.None);
        }

        private static string GetOpenXmlCellText(OfficeXmlNode cell) {
            return string.Join(
                CellParagraphSeparator,
                CellParagraphs(cell)
                    .Select(InnerText)
                    .ToArray());
        }

        private static void EnsureCellHasParagraph(OfficeXmlNode cell) {
            if (cell.Child(W, "p").IsNone) {
                cell.Append(cell.CreateElement(W, "p"));
            }
        }

        private static void RemoveAllChildren(OfficeXmlNode element) {
            foreach (OfficeXmlNode child in element.Children().ToList()) {
                child.Remove();
            }
        }

        private static bool TryParseTableCellLocation(string location, out int tableIndex, out int rowIndex, out int cellIndex) {
            tableIndex = -1;
            rowIndex = -1;
            cellIndex = -1;
            const string tablePrefix = "table[";
            const string rowToken = "]/row[";
            const string cellToken = "]/cell[";
            if (!location.StartsWith(tablePrefix, StringComparison.Ordinal)) {
                return false;
            }

            int rowTokenIndex = location.IndexOf(rowToken, StringComparison.Ordinal);
            int cellTokenIndex = location.IndexOf(cellToken, StringComparison.Ordinal);
            if (rowTokenIndex < 0 || cellTokenIndex < 0 || cellTokenIndex <= rowTokenIndex || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string tableText = location.Substring(tablePrefix.Length, rowTokenIndex - tablePrefix.Length);
            string rowText = location.Substring(rowTokenIndex + rowToken.Length, cellTokenIndex - rowTokenIndex - rowToken.Length);
            string cellText = location.Substring(cellTokenIndex + cellToken.Length, location.Length - cellTokenIndex - cellToken.Length - 1);
            return int.TryParse(tableText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out tableIndex) &&
                   int.TryParse(rowText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out rowIndex) &&
                   int.TryParse(cellText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out cellIndex);
        }

        private static bool TryParseTableRowLocation(string location, out int tableIndex, out int rowIndex) {
            tableIndex = -1;
            rowIndex = -1;
            const string tablePrefix = "table[";
            const string rowToken = "]/row[";
            if (!location.StartsWith(tablePrefix, StringComparison.Ordinal)) {
                return false;
            }

            int rowTokenIndex = location.IndexOf(rowToken, StringComparison.Ordinal);
            if (rowTokenIndex < 0 || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string tableText = location.Substring(tablePrefix.Length, rowTokenIndex - tablePrefix.Length);
            string rowText = location.Substring(rowTokenIndex + rowToken.Length, location.Length - rowTokenIndex - rowToken.Length - 1);
            return int.TryParse(tableText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out tableIndex) &&
                   int.TryParse(rowText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out rowIndex);
        }

        private static string CreateTableCellKey(int tableIndex, int rowIndex, int cellIndex) {
            return tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                   rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                   cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private sealed class RedlineTableEntry {
            internal RedlineTableEntry(string partKey, OfficeXmlNode container, OfficeXmlNode table, int localIndex) {
                PartKey = partKey;
                Container = container;
                Table = table;
                LocalIndex = localIndex;
            }

            internal string PartKey { get; }

            internal OfficeXmlNode Container { get; }

            internal OfficeXmlNode Table { get; }

            internal int LocalIndex { get; }
        }

        /// <summary>A <c>w:ins</c> holding a run with the text, made for <paramref name="context"/> to hold.</summary>
        private static OfficeXmlNode CreateInsertedRun(OfficeXmlNode context, string text, WordComparisonRedlineOptions options) {
            OfficeXmlNode inserted = context.CreateElement(W, "ins");
            OfficeXmlNode run = inserted.CreateElement(W, "r");
            run.SetAttribute(W, "rsidR", WordHeadersAndFooters.GenerateRsid());
            run.Append(CreateTextElement(run, "t", text));
            SetRevisionAttributes(inserted, options.Author, options.DateTime ?? DateTime.Now);
            inserted.Append(run);
            return inserted;
        }

        /// <summary>A <c>w:del</c> holding a run with the deleted text, made for <paramref name="context"/> to hold.</summary>
        private static OfficeXmlNode CreateDeletedRun(OfficeXmlNode context, string text, WordComparisonRedlineOptions options) {
            OfficeXmlNode deleted = context.CreateElement(W, "del");
            OfficeXmlNode run = deleted.CreateElement(W, "r");
            run.SetAttribute(W, "rsidDel", WordHeadersAndFooters.GenerateRsid());
            run.Append(CreateTextElement(run, "delText", text));
            SetRevisionAttributes(deleted, options.Author, options.DateTime ?? DateTime.Now);
            deleted.Append(run);
            return deleted;
        }

        private static OfficeXmlNode CreateTextElement(OfficeXmlNode run, string localName, string text) {
            OfficeXmlNode element = run.CreateElement(W, localName);
            element.SetAttribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Xml, "space", "preserve");
            element.SetInnerText(text);
            return element;
        }

        /// <summary>Sets the author, date and a new revision id of a tracked change, as its typed members did.</summary>
        private static void SetRevisionAttributes(OfficeXmlNode revision, string? author, DateTime date) {
            if (author != null) {
                revision.SetAttribute(W, "author", author);
            }

            revision.SetAttribute(W, "date", WordPackageView.FormatDateTime(date));
            revision.SetAttribute(W, "id", WordHeadersAndFooters.GenerateRevisionId());
        }

        private static void AppendRedlineSummary(WordDocument document, WordComparisonResult result) {
            document.AddParagraph("Summary").SetStyle(WordParagraphStyles.Heading2);
            WordTable table = document.AddTable(4, 2);
            SetCellText(table, 0, 0, "Metric");
            SetCellText(table, 0, 1, "Value");
            SetCellText(table, 1, 0, "Source");
            SetCellText(table, 1, 1, result.SourcePath);
            SetCellText(table, 2, 0, "Target");
            SetCellText(table, 2, 1, result.TargetPath);
            SetCellText(table, 3, 0, "Findings");
            SetCellText(table, 3, 1, result.Findings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private static void AppendRedlineFindingsTable(WordDocument document, WordComparisonResult result) {
            document.AddParagraph("Findings").SetStyle(WordParagraphStyles.Heading2);
            WordTable table = document.AddTable(result.Findings.Count + 1, 6);
            SetCellText(table, 0, 0, "#");
            SetCellText(table, 0, 1, "Scope");
            SetCellText(table, 0, 2, "Change");
            SetCellText(table, 0, 3, "Location");
            SetCellText(table, 0, 4, "Source");
            SetCellText(table, 0, 5, "Target");

            for (int i = 0; i < result.Findings.Count; i++) {
                WordComparisonFinding finding = result.Findings[i];
                int row = i + 1;
                SetCellText(table, row, 0, i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                SetCellText(table, row, 1, finding.Scope.ToString());
                SetCellText(table, row, 2, finding.ChangeKind.ToString());
                SetCellText(table, row, 3, finding.Location);
                SetCellText(table, row, 4, finding.SourceText ?? string.Empty);
                SetCellText(table, row, 5, finding.TargetText ?? string.Empty);
            }
        }

        private static void AppendTrackedRedlineFindings(WordDocument document, WordComparisonResult result, WordComparisonRedlineOptions options) {
            document.AddParagraph("Tracked Changes").SetStyle(WordParagraphStyles.Heading2);
            if (result.Findings.Count == 0) {
                document.AddParagraph("No tracked changes were generated because no structural differences were detected.");
                return;
            }

            bool wroteAnyRevision = false;
            foreach (WordComparisonFinding finding in result.Findings) {
                document.AddParagraph(finding.Location + " - " + finding.Scope + " " + finding.ChangeKind);
                WordParagraph paragraph = document.AddParagraph();
                bool wroteRevision = false;

                if (ShouldTrackFinding(finding, options) && ShouldEmitDeletedText(finding)) {
                    paragraph.AddDeletedText(finding.SourceText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (ShouldTrackFinding(finding, options) && ShouldEmitInsertedText(finding)) {
                    paragraph.AddInsertedText(finding.TargetText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (!wroteRevision) {
                    paragraph.SetText(finding.Message);
                } else {
                    wroteAnyRevision = true;
                }
            }

            if (!wroteAnyRevision) {
                document.AddParagraph("No tracked changes were generated because the selected redline policy kept all findings report-only.");
            }
        }

        private static void AppendInPlaceFeatureAndReviewFindings(WordDocument document, WordComparisonResult result, WordComparisonRedlineOptions options) {
            WordComparisonFinding[] fallbackFindings = result.Findings
                .Where(finding => (IsInPlaceFallbackFeatureFinding(finding) || IsReviewFinding(finding)) && ShouldTrackFinding(finding, options))
                .ToArray();
            if (fallbackFindings.Length == 0) {
                return;
            }

            document.AddParagraph("Tracked Review Changes").SetStyle(WordParagraphStyles.Heading2);
            foreach (WordComparisonFinding finding in fallbackFindings) {
                document.AddParagraph(finding.Location + " - " + finding.Scope + " " + finding.ChangeKind);
                WordParagraph paragraph = document.AddParagraph();
                bool wroteRevision = false;

                if (ShouldEmitDeletedText(finding)) {
                    paragraph.AddDeletedText(finding.SourceText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (ShouldEmitInsertedText(finding)) {
                    paragraph.AddInsertedText(finding.TargetText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (!wroteRevision) {
                    paragraph.SetText(finding.Message);
                }
            }
        }

        private static bool ShouldAppendTrackedRedlineFindings(WordComparisonRedlineOptions options) {
            return options.TrackTextFindings ||
                   options.TrackFeatureFindings ||
                   options.TrackReviewFindings ||
                   options.TrackFormattingFindings;
        }

        private static bool ShouldTrackFinding(WordComparisonFinding finding, WordComparisonRedlineOptions options) {
            if (!options.TrackFeatureFindings && IsFeatureFinding(finding)) {
                return false;
            }

            if (!options.TrackReviewFindings && IsReviewFinding(finding)) {
                return false;
            }

            if (!options.TrackFormattingFindings && IsFormattingFinding(finding)) {
                return false;
            }

            if (!options.TrackTextFindings && IsTextFinding(finding)) {
                return false;
            }

            return true;
        }

        private static bool IsFeatureFinding(WordComparisonFinding finding) {
            return finding.Scope is WordComparisonScope.Field
                or WordComparisonScope.ContentControl
                or WordComparisonScope.Bookmark
                or WordComparisonScope.Hyperlink
                or WordComparisonScope.List
                or WordComparisonScope.Image
                or WordComparisonScope.Shape;
        }

        private static bool IsInPlaceFallbackFeatureFinding(WordComparisonFinding finding) {
            return finding.Scope is WordComparisonScope.Field
                or WordComparisonScope.Bookmark
                or WordComparisonScope.Hyperlink
                or WordComparisonScope.List
                or WordComparisonScope.Shape;
        }

        private static bool IsReviewFinding(WordComparisonFinding finding) {
            return finding.Scope is WordComparisonScope.Comment or WordComparisonScope.Revision;
        }

        private static bool IsFormattingFinding(WordComparisonFinding finding) {
            if (finding.Message.IndexOf("formatting", StringComparison.OrdinalIgnoreCase) >= 0) {
                return true;
            }

            string sourceText = finding.SourceText ?? string.Empty;
            string targetText = finding.TargetText ?? string.Empty;
            return ContainsReviewFormattingKind(sourceText) || ContainsReviewFormattingKind(targetText);
        }

        private static bool IsTextFinding(WordComparisonFinding finding) {
            return HasTrackedText(finding)
                && !IsFeatureFinding(finding)
                && !IsReviewFinding(finding)
                && !IsFormattingFinding(finding);
        }

        private static bool ContainsReviewFormattingKind(string value) {
            return value.IndexOf("Formatting", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ShouldEmitDeletedText(WordComparisonFinding finding) {
            return finding.ChangeKind != WordComparisonChangeKind.Inserted && !string.IsNullOrEmpty(finding.SourceText);
        }

        private static bool ShouldEmitInsertedText(WordComparisonFinding finding) {
            return finding.ChangeKind != WordComparisonChangeKind.Deleted && !string.IsNullOrEmpty(finding.TargetText);
        }

        private static void SetCellText(WordTable table, int row, int column, string text) {
            table.Rows[row].Cells[column].Paragraphs[0].SetText(text);
        }
    }
}
