using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const int LcsCellLimit = 1_000_000;
        private const int MaxComparisonAlignmentWindow = 256;
        private const int BodyPartOrderBase = 0;
        private const int HeaderPartOrderBase = 1_000_000;
        private const int FooterPartOrderBase = 2_000_000;
        private const int FootnotePartOrderBase = 3_000_000;
        private const int EndnotePartOrderBase = 4_000_000;
        private const int RelatedPartOrderStride = 100_000;
        private const string BodyPartKey = "body";
        private const string HeaderPartKeyPrefix = "header:";
        private const string FooterPartKeyPrefix = "footer:";
        private const string FootnotePartKeyPrefix = "footnote:";
        private const string EndnotePartKeyPrefix = "endnote:";
        private const string TableRowSeparator = "\n";
        private const string CellParagraphSeparator = "[ParagraphBreak]";

        /// <summary>
        /// Compares two documents and returns a machine-readable summary of structural differences.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <returns>A deterministic comparison result that can be used for review reports and automation.</returns>
        public static WordComparisonResult CompareStructure(string sourcePath, string targetPath) {
            return CompareStructure(sourcePath, targetPath, null);
        }

        /// <summary>
        /// Compares two documents and returns a machine-readable summary of structural differences.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <param name="options">Optional comparison switches for text normalization and feature families.</param>
        /// <returns>A deterministic comparison result that can be used for review reports and automation.</returns>
        public static WordComparisonResult CompareStructure(string sourcePath, string targetPath, WordComparisonOptions? options) {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));
            options ??= WordComparisonOptions.Default;

            using WordDocument source = WordDocument.Load(sourcePath);
            using WordDocument target = WordDocument.Load(targetPath);

            WordComparisonResult result = new(sourcePath, targetPath);
            AnalyzeParagraphs(source, target, result, options);
            if (options.CompareFields) {
                AnalyzeFields(source, target, result, options);
            }

            if (options.CompareContentControls) {
                AnalyzeContentControls(source, target, result, options);
            }

            if (options.CompareBookmarks) {
                AnalyzeBookmarks(source, target, result, options);
            }

            if (options.CompareHyperlinks) {
                AnalyzeHyperlinks(source, target, result, options);
            }

            if (options.CompareLists) {
                AnalyzeLists(source, target, result, options);
            }

            if (options.CompareComments) {
                AnalyzeComments(source, target, result, options);
            }

            if (options.CompareRevisions) {
                AnalyzeRevisions(source, target, result, options);
            }

            AnalyzeTables(source, target, result, options);
            if (options.CompareImages) {
                AnalyzeImages(source, target, result);
            }

            if (options.CompareBlockOrder) {
                AnalyzeBlockOrder(source, target, result, options);
            }

            ApplyScopeFilters(result, options);
            result.SortFindingsByDocumentOrder();
            return result;
        }

        private static void ApplyScopeFilters(WordComparisonResult result, WordComparisonOptions options) {
            bool hasIncludedScopes = options.IncludedScopes != null && options.IncludedScopes.Count > 0;
            bool hasExcludedScopes = options.ExcludedScopes != null && options.ExcludedScopes.Count > 0;
            if (!hasIncludedScopes && !hasExcludedScopes) {
                return;
            }

            result.RemoveWhere(finding =>
                (hasIncludedScopes && !options.IncludedScopes!.Contains(finding.Scope)) ||
                (hasExcludedScopes && options.ExcludedScopes!.Contains(finding.Scope)));
        }

        private static void AnalyzeTables(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            List<TableSnapshot> sourceTables = GetTableSnapshots(source, options);
            List<TableSnapshot> targetTables = GetTableSnapshots(target, options);
            IReadOnlyList<MatchedIndexPair> matchedTables = FindMatchingIndexes(
                sourceTables.Select(table => table.MatchKey).ToList(),
                targetTables.Select(table => table.MatchKey).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedTables) {
                AddTableRangeFindings(sourceTables, targetTables, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRangeFindings(sourceTables, targetTables, sourceStart, sourceTables.Count, targetStart, targetTables.Count, result, options);
        }

        private static void AddTableRangeFindings(
            IReadOnlyList<TableSnapshot> sourceTables,
            IReadOnlyList<TableSnapshot> targetTables,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex + 1], options) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex], options)) {
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex + 1], targetTables[targetIndex], options) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex], options)) {
                    AddDeletedTableFinding(sourceTables, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                if (!string.Equals(sourceTables[sourceIndex].PartKey, targetTables[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedTableFinding(sourceTables, sourceIndex, result);
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                AnalyzeTable(sourceTables[sourceIndex], targetTables[targetIndex], targetIndex, targetTables[targetIndex].DocumentOrder, result, options);
                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableFinding(targetTables, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableFinding(sourceTables, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedTableFinding(IReadOnlyList<TableSnapshot> targetTables, int tableIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Table,
                WordComparisonChangeKind.Inserted,
                TableLocation(tableIndex),
                null,
                tableIndex,
                null,
                targetTables[tableIndex].Text,
                "Table inserted."),
                targetTables[tableIndex].DocumentOrder);
        }

        private static void AddDeletedTableFinding(IReadOnlyList<TableSnapshot> sourceTables, int tableIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Table,
                WordComparisonChangeKind.Deleted,
                TableLocation(tableIndex),
                tableIndex,
                null,
                sourceTables[tableIndex].Text,
                null,
                "Table deleted."),
                sourceTables[tableIndex].DocumentOrder);
        }

        private static void AnalyzeTable(TableSnapshot source, TableSnapshot target, int tableIndex, int tableDocumentOrder, WordComparisonResult result, WordComparisonOptions options) {
            List<WordTableRow> sourceRows = source.Table.Rows.ToList();
            List<WordTableRow> targetRows = target.Table.Rows.ToList();
            IReadOnlyList<MatchedIndexPair> matchedRows = FindMatchingIndexes(
                sourceRows.Select(row => GetRowMatchKey(row, source.Part, options)).ToList(),
                targetRows.Select(row => GetRowMatchKey(row, target.Part, options)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedRows) {
                AddTableRowRangeFindings(sourceRows, targetRows, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRowRangeFindings(sourceRows, targetRows, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, sourceRows.Count, targetStart, targetRows.Count, result, options);
        }

        private static void AddTableRowRangeFindings(
            IReadOnlyList<WordTableRow> sourceRows,
            IReadOnlyList<WordTableRow> targetRows,
            OpenXmlPart? sourcePart,
            OpenXmlPart? targetPart,
            int tableIndex,
            int tableDocumentOrder,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                int betterTargetIndex = FindBetterTargetRowAlignmentIndex(sourceRows[sourceIndex], targetRows, targetIndex, targetEnd, options);
                if (betterTargetIndex > targetIndex) {
                    while (targetIndex < betterTargetIndex) {
                        AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                        targetIndex++;
                    }

                    continue;
                }

                int betterSourceIndex = FindBetterSourceRowAlignmentIndex(sourceRows, sourceIndex, sourceEnd, targetRows[targetIndex], options);
                if (betterSourceIndex > sourceIndex) {
                    while (sourceIndex < betterSourceIndex) {
                        AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                        sourceIndex++;
                    }

                    continue;
                }

                if (sourceRows[sourceIndex].Cells.Count != targetRows[targetIndex].Cells.Count &&
                    AreComparisonTextEqual(GetRowText(sourceRows[sourceIndex]), GetRowText(targetRows[targetIndex]), options)) {
                    AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                    AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                AnalyzeTableRow(sourceRows[sourceIndex], targetRows[targetIndex], sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceIndex, targetIndex, result, options);
                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static int FindBetterTargetRowAlignmentIndex(
            WordTableRow sourceRow,
            IReadOnlyList<WordTableRow> targetRows,
            int targetStart,
            int targetEnd,
            WordComparisonOptions options) {
            double currentSimilarity = GetRowSimilarity(sourceRow, targetRows[targetStart], options);
            int bestIndex = targetStart;
            double bestSimilarity = currentSimilarity;

            int boundedTargetEnd = Math.Min(targetEnd, targetStart + MaxComparisonAlignmentWindow);
            for (int index = targetStart + 1; index < boundedTargetEnd; index++) {
                double similarity = GetRowSimilarity(sourceRow, targetRows[index], options);
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestSimilarity > currentSimilarity ? bestIndex : targetStart;
        }

        private static int FindBetterSourceRowAlignmentIndex(
            IReadOnlyList<WordTableRow> sourceRows,
            int sourceStart,
            int sourceEnd,
            WordTableRow targetRow,
            WordComparisonOptions options) {
            double currentSimilarity = GetRowSimilarity(sourceRows[sourceStart], targetRow, options);
            int bestIndex = sourceStart;
            double bestSimilarity = currentSimilarity;

            int boundedSourceEnd = Math.Min(sourceEnd, sourceStart + MaxComparisonAlignmentWindow);
            for (int index = sourceStart + 1; index < boundedSourceEnd; index++) {
                double similarity = GetRowSimilarity(sourceRows[index], targetRow, options);
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestSimilarity > currentSimilarity ? bestIndex : sourceStart;
        }

        private static void AddInsertedTableRowFinding(IReadOnlyList<WordTableRow> targetRows, int tableIndex, int tableDocumentOrder, int rowIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableRow,
                WordComparisonChangeKind.Inserted,
                RowLocation(tableIndex, rowIndex),
                null,
                rowIndex,
                null,
                GetRowText(targetRows[rowIndex]),
                "Table row inserted."),
                GetTableChildDocumentOrder(tableDocumentOrder, targetRows[rowIndex]._tableRow));
        }

        private static void AddDeletedTableRowFinding(IReadOnlyList<WordTableRow> sourceRows, int tableIndex, int tableDocumentOrder, int rowIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableRow,
                WordComparisonChangeKind.Deleted,
                RowLocation(tableIndex, rowIndex),
                rowIndex,
                null,
                GetRowText(sourceRows[rowIndex]),
                null,
                "Table row deleted."),
                GetTableChildDocumentOrder(tableDocumentOrder, sourceRows[rowIndex]._tableRow));
        }

        private static void AnalyzeTableRow(WordTableRow source, WordTableRow target, OpenXmlPart? sourcePart, OpenXmlPart? targetPart, int tableIndex, int tableDocumentOrder, int sourceRowIndex, int targetRowIndex, WordComparisonResult result, WordComparisonOptions options) {
            List<WordTableCell> sourceCells = source.Cells.ToList();
            List<WordTableCell> targetCells = target.Cells.ToList();
            IReadOnlyList<MatchedIndexPair> matchedCells = FindMatchingIndexes(
                sourceCells.Select(cell => GetCellMatchKey(cell, sourcePart, options)).ToList(),
                targetCells.Select(cell => GetCellMatchKey(cell, targetPart, options)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedCells) {
                AddTableCellRangeFindings(sourceCells, targetCells, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableCellRangeFindings(sourceCells, targetCells, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, sourceCells.Count, targetStart, targetCells.Count, result, options);
        }

        private static void AddTableCellRangeFindings(
            IReadOnlyList<WordTableCell> sourceCells,
            IReadOnlyList<WordTableCell> targetCells,
            OpenXmlPart? sourcePart,
            OpenXmlPart? targetPart,
            int tableIndex,
            int tableDocumentOrder,
            int sourceRowIndex,
            int targetRowIndex,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex + 1], options) >
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex], options)) {
                    AddInsertedTableCellFinding(targetCells, tableIndex, tableDocumentOrder, targetRowIndex, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetCellSimilarity(sourceCells[sourceIndex + 1], targetCells[targetIndex], options) >
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex], options)) {
                    AddDeletedTableCellFinding(sourceCells, tableIndex, tableDocumentOrder, sourceRowIndex, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                string sourceText = GetCellText(sourceCells[sourceIndex]);
                string targetText = GetCellText(targetCells[targetIndex]);
                string sourceMatchText = GetCellMatchText(sourceCells[sourceIndex], sourcePart, options);
                string targetMatchText = GetCellMatchText(targetCells[targetIndex], targetPart, options);
                string sourceCellShape = GetCellShape(sourceCells[sourceIndex]);
                string targetCellShape = GetCellShape(targetCells[targetIndex]);

                if (!string.Equals(sourceMatchText, targetMatchText, StringComparison.Ordinal) ||
                    !string.Equals(sourceCellShape, targetCellShape, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.TableCell,
                        WordComparisonChangeKind.Modified,
                        CellLocation(tableIndex, targetRowIndex, targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        AreComparisonTextEqual(sourceText, targetText, options) ? "Table cell structure changed." : "Table cell text changed."),
                        GetTableChildDocumentOrder(tableDocumentOrder, targetCells[targetIndex]._tableCell));
                }

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableCellFinding(targetCells, tableIndex, tableDocumentOrder, targetRowIndex, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableCellFinding(sourceCells, tableIndex, tableDocumentOrder, sourceRowIndex, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedTableCellFinding(IReadOnlyList<WordTableCell> targetCells, int tableIndex, int tableDocumentOrder, int targetRowIndex, int cellIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableCell,
                WordComparisonChangeKind.Inserted,
                CellLocation(tableIndex, targetRowIndex, cellIndex),
                null,
                cellIndex,
                null,
                GetCellText(targetCells[cellIndex]),
                "Table cell inserted."),
                GetTableChildDocumentOrder(tableDocumentOrder, targetCells[cellIndex]._tableCell));
        }

        private static void AddDeletedTableCellFinding(IReadOnlyList<WordTableCell> sourceCells, int tableIndex, int tableDocumentOrder, int sourceRowIndex, int cellIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableCell,
                WordComparisonChangeKind.Deleted,
                CellLocation(tableIndex, sourceRowIndex, cellIndex),
                cellIndex,
                null,
                GetCellText(sourceCells[cellIndex]),
                null,
                "Table cell deleted."),
                GetTableChildDocumentOrder(tableDocumentOrder, sourceCells[cellIndex]._tableCell));
        }

        private static IReadOnlyList<MatchedIndexPair> FindMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) where T : notnull {
            if ((long)(source.Count + 1) * (target.Count + 1) > LcsCellLimit) {
                return FindGreedyMatchingIndexes(source, target, comparer);
            }

            int[,] lengths = new int[source.Count + 1, target.Count + 1];

            for (int sourceIndex = source.Count - 1; sourceIndex >= 0; sourceIndex--) {
                for (int targetIndex = target.Count - 1; targetIndex >= 0; targetIndex--) {
                    lengths[sourceIndex, targetIndex] = comparer.Equals(source[sourceIndex], target[targetIndex])
                        ? lengths[sourceIndex + 1, targetIndex + 1] + 1
                        : Math.Max(lengths[sourceIndex + 1, targetIndex], lengths[sourceIndex, targetIndex + 1]);
                }
            }

            var matches = new List<MatchedIndexPair>();
            int sourceCursor = 0;
            int targetCursor = 0;

            while (sourceCursor < source.Count && targetCursor < target.Count) {
                if (comparer.Equals(source[sourceCursor], target[targetCursor])) {
                    matches.Add(new MatchedIndexPair(sourceCursor, targetCursor));
                    sourceCursor++;
                    targetCursor++;
                } else if (lengths[sourceCursor + 1, targetCursor] >= lengths[sourceCursor, targetCursor + 1]) {
                    sourceCursor++;
                } else {
                    targetCursor++;
                }
            }

            return matches;
        }

        private static IReadOnlyList<MatchedIndexPair> FindGreedyMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) where T : notnull {
            var matches = new List<MatchedIndexPair>();
            var targetIndexes = new Dictionary<T, Queue<int>>(comparer);
            for (int targetIndex = 0; targetIndex < target.Count; targetIndex++) {
                if (!targetIndexes.TryGetValue(target[targetIndex], out Queue<int>? indexes)) {
                    indexes = new Queue<int>();
                    targetIndexes.Add(target[targetIndex], indexes);
                }

                indexes.Enqueue(targetIndex);
            }

            int targetCursor = 0;

            for (int sourceIndex = 0; sourceIndex < source.Count && targetCursor < target.Count; sourceIndex++) {
                if (!targetIndexes.TryGetValue(source[sourceIndex], out Queue<int>? indexes)) {
                    continue;
                }

                while (indexes.Count > 0 && indexes.Peek() < targetCursor) {
                    indexes.Dequeue();
                }

                if (indexes.Count > 0) {
                    int targetIndex = indexes.Dequeue();
                    matches.Add(new MatchedIndexPair(sourceIndex, targetIndex));
                    targetCursor = targetIndex + 1;
                }
            }

            return matches;
        }

        private static List<TableSnapshot> GetTableSnapshots(WordDocument document, WordComparisonOptions options) {
            var snapshots = new List<TableSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddTableSnapshots(snapshots, document, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddTableSnapshots(snapshots, document, headerPartKey.Key, headerPartKey.Key.Header, headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddTableSnapshots(snapshots, document, footerPartKey.Key, footerPartKey.Key.Footer, footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    AddTableSnapshots(snapshots, document, mainPart.FootnotesPart, footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    AddTableSnapshots(snapshots, document, mainPart.EndnotesPart, endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddTableSnapshots(List<TableSnapshot> snapshots, WordDocument document, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Table table) {
                    continue;
                }

                var wordTable = new WordTable(document, table);
                string text = GetTableText(wordTable);
                string matchText = GetTableMatchText(wordTable, part, options);
                snapshots.Add(new TableSnapshot(wordTable, part, text, matchText, GetTableMatchKey(partKey, wordTable, part, options), partKey, ordered.DocumentOrder));
            }
        }

        private static string GetTableText(WordTable table) {
            return string.Join(TableRowSeparator, table.Rows.Select(GetRowText).ToArray());
        }

        private static string GetRowText(WordTableRow row) {
            return string.Join(" | ", row.Cells.Select(GetCellText).ToArray());
        }

        private static string GetTableMatchText(WordTable table, OpenXmlPart? part, WordComparisonOptions options) {
            return string.Join(TableRowSeparator, table.Rows.Select(row => GetRowMatchText(row, part, options)).ToArray());
        }

        private static string GetRowMatchText(WordTableRow row, OpenXmlPart? part, WordComparisonOptions options) {
            return string.Join(" | ", row.Cells.Select(cell => GetCellMatchText(cell, part, options)).ToArray());
        }

        private static string GetTableMatchKey(string partKey, WordTable table, OpenXmlPart? part, WordComparisonOptions options) {
            return partKey + TableRowSeparator + GetTableShape(table) + TableRowSeparator + string.Join(TableRowSeparator, table.Rows.Select(row => GetRowMatchKey(row, part, options)).ToArray());
        }

        private static string GetTableShape(WordTable table) {
            return string.Join(";", table.Rows.Select(GetRowShape).ToArray());
        }

        private static string GetRowMatchKey(WordTableRow row, OpenXmlPart? part, WordComparisonOptions options) {
            return GetRowShape(row) + TableRowSeparator + string.Join(TableRowSeparator, row.Cells.Select(cell => EncodeMatchText(GetCellMatchText(cell, part, options))).ToArray());
        }

        private static string GetRowShape(WordTableRow row) {
            return row.Cells.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                   string.Join(",", row.Cells.Select(GetCellShape).ToArray());
        }

        private static string GetCellShape(WordTableCell cell) {
            TableCellProperties? properties = cell._tableCell.GetFirstChild<TableCellProperties>();
            if (properties == null) {
                return string.Empty;
            }

            string gridSpan = properties.GridSpan?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            string horizontalMerge = GetMergeShapeValue(properties.HorizontalMerge);
            string verticalMerge = GetMergeShapeValue(properties.VerticalMerge);
            return "span=" + gridSpan + ";h=" + horizontalMerge + ";v=" + verticalMerge;
        }

        private static string GetMergeShapeValue(OpenXmlLeafElement? merge) {
            if (merge == null) {
                return string.Empty;
            }

            if (merge is HorizontalMerge horizontalMerge) {
                return horizontalMerge.Val?.Value.ToString() ?? "continue";
            }

            if (merge is VerticalMerge verticalMerge) {
                return verticalMerge.Val?.Value.ToString() ?? "continue";
            }

            return string.Empty;
        }

        private static string GetCellText(WordTableCell cell) {
            return string.Join(
                CellParagraphSeparator,
                cell._tableCell.Descendants<Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<TableCell>().FirstOrDefault(), cell._tableCell))
                    .Select(GetParagraphText)
                    .ToArray());
        }

        private static string GetCellMatchText(WordTableCell cell) {
            return GetCellMatchText(cell, null);
        }

        private static string GetCellMatchText(WordTableCell cell, OpenXmlPart? part) {
            return string.Join(
                CellParagraphSeparator,
                cell._tableCell.Descendants<Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<TableCell>().FirstOrDefault(), cell._tableCell))
                    .Select(paragraph => GetParagraphMatchText(paragraph, part))
                    .ToArray());
        }

        private static string GetCellMatchText(WordTableCell cell, OpenXmlPart? part, WordComparisonOptions options) {
            return string.Join(
                CellParagraphSeparator,
                cell._tableCell.Descendants<Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<TableCell>().FirstOrDefault(), cell._tableCell))
                    .Select(paragraph => GetParagraphMatchText(paragraph, part, options))
                    .ToArray());
        }

        private static string GetCellMatchKey(WordTableCell cell, OpenXmlPart? part) {
            return GetCellShape(cell) + CellParagraphSeparator + GetCellMatchText(cell, part);
        }

        private static string GetCellMatchKey(WordTableCell cell, OpenXmlPart? part, WordComparisonOptions options) {
            return GetCellShape(cell) + CellParagraphSeparator + GetCellMatchText(cell, part, options);
        }

        private static string EncodeMatchText(string value) {
            return value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + value;
        }

        private static double GetRowSimilarity(WordTableRow source, WordTableRow target, WordComparisonOptions options) {
            if (source.Cells.Count != target.Cells.Count) {
                return 0;
            }

            return GetContainmentAwareTextSimilarity(NormalizeComparisonText(GetRowText(source), options), NormalizeComparisonText(GetRowText(target), options));
        }

        private static double GetTableSimilarity(TableSnapshot source, TableSnapshot target, WordComparisonOptions options) {
            if (!string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) {
                return 0;
            }

            return GetTextSimilarity(NormalizeComparisonText(source.Text, options), NormalizeComparisonText(target.Text, options));
        }

        private static double GetCellSimilarity(WordTableCell source, WordTableCell target, WordComparisonOptions options) {
            return GetTextSimilarity(NormalizeComparisonText(GetCellText(source), options), NormalizeComparisonText(GetCellText(target), options));
        }

        private static string GetStableElementPath(OpenXmlElement element) {
            var segments = new Stack<string>();
            OpenXmlElement? current = element;
            while (current != null && current.Parent != null) {
                OpenXmlElement parent = current.Parent;
                int ordinal = parent.Elements()
                    .Where(item => item.GetType() == current.GetType())
                    .TakeWhile(item => !ReferenceEquals(item, current))
                    .Count();
                segments.Push(current.GetType().Name + "[" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
                current = parent;
            }

            return string.Join("/", segments.ToArray());
        }

        private static IEnumerable<OrderedElement> EnumerateDescendantsWithOrder(OpenXmlElement? container, int orderBase) {
            if (container == null) {
                yield break;
            }

            int order = orderBase;
            foreach (OpenXmlElement element in EnumerateComparableDescendants(container)) {
                yield return new OrderedElement(element, order);
                order++;
            }
        }

        private static IEnumerable<OpenXmlElement> EnumerateComparableDescendants(OpenXmlElement container) {
            foreach (OpenXmlElement child in GetComparableChildren(container)) {
                yield return child;
                foreach (OpenXmlElement descendant in EnumerateComparableDescendants(child)) {
                    yield return descendant;
                }
            }
        }

        private static IEnumerable<OpenXmlElement> GetComparableChildren(OpenXmlElement element) {
            if (element is AlternateContent alternateContent) {
                foreach (OpenXmlElement child in GetActiveAlternateContentChildren(alternateContent)) {
                    yield return child;
                }

                yield break;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                yield return child;
            }
        }

        private static IEnumerable<OpenXmlElement> GetActiveAlternateContentChildren(AlternateContent alternateContent) {
            AlternateContentChoice? choice = alternateContent.Elements<AlternateContentChoice>().FirstOrDefault();
            if (choice != null) {
                foreach (OpenXmlElement child in choice.ChildElements) {
                    yield return child;
                }

                yield break;
            }

            AlternateContentFallback? fallback = alternateContent.GetFirstChild<AlternateContentFallback>();
            if (fallback == null) {
                yield break;
            }

            foreach (OpenXmlElement child in fallback.ChildElements) {
                yield return child;
            }
        }

        private static string TableLocation(int tableIndex) {
            return "table[" + tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string RowLocation(int tableIndex, int rowIndex) {
            return TableLocation(tableIndex) + "/row[" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string CellLocation(int tableIndex, int rowIndex, int cellIndex) {
            return RowLocation(tableIndex, rowIndex) + "/cell[" + cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static int GetTableChildDocumentOrder(int tableDocumentOrder, OpenXmlElement child) {
            Table? table = child as Table ?? child.Ancestors<Table>().FirstOrDefault();
            if (table == null) {
                return tableDocumentOrder;
            }

            int offset = 1;
            foreach (OpenXmlElement descendant in table.Descendants()) {
                if (ReferenceEquals(descendant, child)) {
                    return tableDocumentOrder + offset;
                }

                offset++;
            }

            return tableDocumentOrder;
        }

        private readonly struct OrderedElement {
            internal OrderedElement(OpenXmlElement element, int documentOrder) {
                Element = element;
                DocumentOrder = documentOrder;
            }

            internal OpenXmlElement Element { get; }

            internal int DocumentOrder { get; }
        }

        private readonly struct MatchedIndexPair {
            internal MatchedIndexPair(int sourceIndex, int targetIndex) {
                SourceIndex = sourceIndex;
                TargetIndex = targetIndex;
            }

            internal int SourceIndex { get; }

            internal int TargetIndex { get; }
        }

        private sealed class TableSnapshot {
            internal TableSnapshot(WordTable table, OpenXmlPart? part, string text, string matchText, string matchKey, string partKey, int documentOrder) {
                Table = table;
                Part = part;
                Text = text;
                MatchText = matchText;
                MatchKey = matchKey;
                PartKey = partKey;
                DocumentOrder = documentOrder;
            }

            internal WordTable Table { get; }

            internal OpenXmlPart? Part { get; }

            internal string Text { get; }

            internal string MatchText { get; }

            internal string MatchKey { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }
        }

    }
}
