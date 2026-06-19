using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const int LcsCellLimit = 1_000_000;
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
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));

            using WordDocument source = WordDocument.Load(sourcePath);
            using WordDocument target = WordDocument.Load(targetPath);

            WordComparisonResult result = new(sourcePath, targetPath);
            AnalyzeParagraphs(source, target, result);
            AnalyzeTables(source, target, result);
            AnalyzeImages(source, target, result);
            AnalyzeBlockOrder(source, target, result);
            result.SortFindingsByDocumentOrder();
            return result;
        }

        private static void AnalyzeTables(WordDocument source, WordDocument target, WordComparisonResult result) {
            List<TableSnapshot> sourceTables = GetTableSnapshots(source);
            List<TableSnapshot> targetTables = GetTableSnapshots(target);
            IReadOnlyList<MatchedIndexPair> matchedTables = FindMatchingIndexes(
                sourceTables.Select(table => table.MatchKey).ToList(),
                targetTables.Select(table => table.MatchKey).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedTables) {
                AddTableRangeFindings(sourceTables, targetTables, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRangeFindings(sourceTables, targetTables, sourceStart, sourceTables.Count, targetStart, targetTables.Count, result);
        }

        private static void AddTableRangeFindings(
            IReadOnlyList<TableSnapshot> sourceTables,
            IReadOnlyList<TableSnapshot> targetTables,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex + 1]) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex])) {
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex + 1], targetTables[targetIndex]) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex])) {
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

                AnalyzeTable(sourceTables[sourceIndex], targetTables[targetIndex], targetIndex, targetTables[targetIndex].DocumentOrder, result);
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

        private static void AnalyzeTable(TableSnapshot source, TableSnapshot target, int tableIndex, int tableDocumentOrder, WordComparisonResult result) {
            List<WordTableRow> sourceRows = source.Table.Rows.ToList();
            List<WordTableRow> targetRows = target.Table.Rows.ToList();
            IReadOnlyList<MatchedIndexPair> matchedRows = FindMatchingIndexes(
                sourceRows.Select(row => GetRowMatchKey(row, source.Part)).ToList(),
                targetRows.Select(row => GetRowMatchKey(row, target.Part)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedRows) {
                AddTableRowRangeFindings(sourceRows, targetRows, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRowRangeFindings(sourceRows, targetRows, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, sourceRows.Count, targetStart, targetRows.Count, result);
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
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetRowSimilarity(sourceRows[sourceIndex], targetRows[targetIndex + 1]) >
                    GetRowSimilarity(sourceRows[sourceIndex], targetRows[targetIndex])) {
                    AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetRowSimilarity(sourceRows[sourceIndex + 1], targetRows[targetIndex]) >
                    GetRowSimilarity(sourceRows[sourceIndex], targetRows[targetIndex])) {
                    AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                if (sourceRows[sourceIndex].Cells.Count != targetRows[targetIndex].Cells.Count &&
                    string.Equals(GetRowText(sourceRows[sourceIndex]), GetRowText(targetRows[targetIndex]), StringComparison.Ordinal)) {
                    AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                    AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                AnalyzeTableRow(sourceRows[sourceIndex], targetRows[targetIndex], sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceIndex, targetIndex, result);
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

        private static void AnalyzeTableRow(WordTableRow source, WordTableRow target, OpenXmlPart? sourcePart, OpenXmlPart? targetPart, int tableIndex, int tableDocumentOrder, int sourceRowIndex, int targetRowIndex, WordComparisonResult result) {
            List<WordTableCell> sourceCells = source.Cells.ToList();
            List<WordTableCell> targetCells = target.Cells.ToList();
            IReadOnlyList<MatchedIndexPair> matchedCells = FindMatchingIndexes(
                sourceCells.Select(cell => GetCellMatchKey(cell, sourcePart)).ToList(),
                targetCells.Select(cell => GetCellMatchKey(cell, targetPart)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedCells) {
                AddTableCellRangeFindings(sourceCells, targetCells, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableCellRangeFindings(sourceCells, targetCells, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, sourceCells.Count, targetStart, targetCells.Count, result);
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
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex + 1]) >
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex])) {
                    AddInsertedTableCellFinding(targetCells, tableIndex, tableDocumentOrder, targetRowIndex, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetCellSimilarity(sourceCells[sourceIndex + 1], targetCells[targetIndex]) >
                    GetCellSimilarity(sourceCells[sourceIndex], targetCells[targetIndex])) {
                    AddDeletedTableCellFinding(sourceCells, tableIndex, tableDocumentOrder, sourceRowIndex, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                string sourceText = GetCellText(sourceCells[sourceIndex]);
                string targetText = GetCellText(targetCells[targetIndex]);
                string sourceMatchText = GetCellMatchText(sourceCells[sourceIndex], sourcePart);
                string targetMatchText = GetCellMatchText(targetCells[targetIndex], targetPart);
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
                        string.Equals(sourceText, targetText, StringComparison.Ordinal) ? "Table cell structure changed." : "Table cell text changed."),
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

        private static void AnalyzeImages(WordDocument source, WordDocument target, WordComparisonResult result) {
            IReadOnlyList<ImageSnapshot> sourceImages = GetImageSnapshots(source);
            IReadOnlyList<ImageSnapshot> targetImages = GetImageSnapshots(target);
            IReadOnlyList<MatchedIndexPair> matchedImages = FindMatchingIndexes(
                sourceImages,
                targetImages,
                ImageSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedImages) {
                AddImageRangeFindings(sourceImages, targetImages, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddImageRangeFindings(sourceImages, targetImages, sourceStart, sourceImages.Count, targetStart, targetImages.Count, result);
            AddImagePositionFindings(sourceImages, targetImages, matchedImages, result);
        }

        private static void AddImagePositionFindings(
            IReadOnlyList<ImageSnapshot> sourceImages,
            IReadOnlyList<ImageSnapshot> targetImages,
            IReadOnlyList<MatchedIndexPair> matchedImages,
            WordComparisonResult result) {
            if (sourceImages.Count != targetImages.Count || matchedImages.Count != sourceImages.Count) {
                return;
            }

            foreach (MatchedIndexPair match in matchedImages) {
                ImageSnapshot sourceImage = sourceImages[match.SourceIndex];
                ImageSnapshot targetImage = targetImages[match.TargetIndex];
                if (!ImageSnapshotEqualityComparer.Instance.Equals(sourceImage, targetImage) ||
                    string.Equals(sourceImage.PositionKey, targetImage.PositionKey, StringComparison.Ordinal)) {
                    continue;
                }

                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Modified,
                    ImageLocation(match.TargetIndex),
                    match.SourceIndex,
                    match.TargetIndex,
                    sourceImage.DisplayText,
                    targetImage.DisplayText,
                    "Image position changed."),
                    targetImage.DocumentOrder);
            }
        }

        private static void AddImageRangeFindings(
            IReadOnlyList<ImageSnapshot> sourceImages,
            IReadOnlyList<ImageSnapshot> targetImages,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex + 1]) >
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex])) {
                    AddInsertedImageFinding(targetImages, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetImageSimilarity(sourceImages[sourceIndex + 1], targetImages[targetIndex]) >
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex])) {
                    AddDeletedImageFinding(sourceImages, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                if (!string.Equals(sourceImages[sourceIndex].PartKey, targetImages[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedImageFinding(sourceImages, sourceIndex, result);
                    AddInsertedImageFinding(targetImages, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Modified,
                    ImageLocation(targetIndex),
                    sourceIndex,
                    targetIndex,
                    sourceImages[sourceIndex].DisplayText,
                    targetImages[targetIndex].DisplayText,
                    "Image payload changed."),
                    targetImages[targetIndex].DocumentOrder);

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedImageFinding(targetImages, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedImageFinding(sourceImages, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedImageFinding(IReadOnlyList<ImageSnapshot> targetImages, int imageIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Image,
                WordComparisonChangeKind.Inserted,
                ImageLocation(imageIndex),
                null,
                imageIndex,
                null,
                targetImages[imageIndex].DisplayText,
                "Image inserted."),
                targetImages[imageIndex].DocumentOrder);
        }

        private static void AddDeletedImageFinding(IReadOnlyList<ImageSnapshot> sourceImages, int imageIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Image,
                WordComparisonChangeKind.Deleted,
                ImageLocation(imageIndex),
                imageIndex,
                null,
                sourceImages[imageIndex].DisplayText,
                null,
                "Image deleted."),
                sourceImages[imageIndex].DocumentOrder);
        }

        private static IReadOnlyList<MatchedIndexPair> FindMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) {
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

        private static IReadOnlyList<MatchedIndexPair> FindGreedyMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) {
            var matches = new List<MatchedIndexPair>();
            int targetCursor = 0;

            for (int sourceIndex = 0; sourceIndex < source.Count && targetCursor < target.Count; sourceIndex++) {
                for (int targetIndex = targetCursor; targetIndex < target.Count; targetIndex++) {
                    if (!comparer.Equals(source[sourceIndex], target[targetIndex])) {
                        continue;
                    }

                    matches.Add(new MatchedIndexPair(sourceIndex, targetIndex));
                    targetCursor = targetIndex + 1;
                    break;
                }
            }

            return matches;
        }

        private static List<TableSnapshot> GetTableSnapshots(WordDocument document) {
            var snapshots = new List<TableSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddTableSnapshots(snapshots, document, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                Dictionary<HeaderPart, string> headerPartKeys = CreateHeaderPartKeys(mainPart);
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (!headerPartKeys.TryGetValue(headerPart, out string? headerPartKey)) {
                        continue;
                    }

                    AddTableSnapshots(snapshots, document, headerPart, headerPart.Header, headerPartKey, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                Dictionary<FooterPart, string> footerPartKeys = CreateFooterPartKeys(mainPart);
                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (!footerPartKeys.TryGetValue(footerPart, out string? footerPartKey)) {
                        continue;
                    }

                    AddTableSnapshots(snapshots, document, footerPart, footerPart.Footer, footerPartKey, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                int footnoteIndex = 0;
                foreach (Footnote footnote in mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>()) {
                    if (!IsVisibleNote(footnote)) {
                        continue;
                    }

                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddTableSnapshots(snapshots, document, mainPart.FootnotesPart, footnote, FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                    footnoteIndex++;
                }

                int endnoteIndex = 0;
                foreach (Endnote endnote in mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? Enumerable.Empty<Endnote>()) {
                    if (!IsVisibleNote(endnote)) {
                        continue;
                    }

                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddTableSnapshots(snapshots, document, mainPart.EndnotesPart, endnote, EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                    endnoteIndex++;
                }
            }

            return snapshots;
        }

        private static void AddTableSnapshots(List<TableSnapshot> snapshots, WordDocument document, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Table table) {
                    continue;
                }

                var wordTable = new WordTable(document, table);
                string text = GetTableText(wordTable);
                string matchText = GetTableMatchText(wordTable, part);
                snapshots.Add(new TableSnapshot(wordTable, part, text, matchText, GetTableMatchKey(partKey, wordTable, part), partKey, ordered.DocumentOrder));
            }
        }

        private static string GetTableText(WordTable table) {
            return string.Join(TableRowSeparator, table.Rows.Select(GetRowText).ToArray());
        }

        private static string GetRowText(WordTableRow row) {
            return string.Join(" | ", row.Cells.Select(GetCellText).ToArray());
        }

        private static string GetTableMatchText(WordTable table, OpenXmlPart? part) {
            return string.Join(TableRowSeparator, table.Rows.Select(row => GetRowMatchText(row, part)).ToArray());
        }

        private static string GetRowMatchText(WordTableRow row, OpenXmlPart? part) {
            return string.Join(" | ", row.Cells.Select(cell => GetCellMatchText(cell, part)).ToArray());
        }

        private static string GetTableMatchKey(string partKey, WordTable table, OpenXmlPart? part) {
            return partKey + TableRowSeparator + GetTableShape(table) + TableRowSeparator + string.Join(TableRowSeparator, table.Rows.Select(row => GetRowMatchKey(row, part)).ToArray());
        }

        private static string GetTableShape(WordTable table) {
            return string.Join(";", table.Rows.Select(GetRowShape).ToArray());
        }

        private static string GetRowMatchKey(WordTableRow row, OpenXmlPart? part) {
            return GetRowShape(row) + TableRowSeparator + string.Join(TableRowSeparator, row.Cells.Select(cell => EncodeMatchText(GetCellMatchText(cell, part))).ToArray());
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
            string horizontalMerge = properties.HorizontalMerge?.Val?.Value.ToString() ?? string.Empty;
            string verticalMerge = properties.VerticalMerge?.Val?.Value.ToString() ?? string.Empty;
            return "span=" + gridSpan + ";h=" + horizontalMerge + ";v=" + verticalMerge;
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

        private static string GetCellMatchKey(WordTableCell cell, OpenXmlPart? part) {
            return GetCellShape(cell) + CellParagraphSeparator + GetCellMatchText(cell, part);
        }

        private static string EncodeMatchText(string value) {
            return value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + value;
        }

        private static double GetRowSimilarity(WordTableRow source, WordTableRow target) {
            if (source.Cells.Count != target.Cells.Count) {
                return 0;
            }

            return GetTextSimilarity(GetRowText(source), GetRowText(target));
        }

        private static double GetTableSimilarity(TableSnapshot source, TableSnapshot target) {
            if (!string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) {
                return 0;
            }

            return GetTextSimilarity(source.Text, target.Text);
        }

        private static double GetCellSimilarity(WordTableCell source, WordTableCell target) {
            return GetTextSimilarity(GetCellText(source), GetCellText(target));
        }

        private static List<ImageSnapshot> GetImageSnapshots(WordDocument document) {
            var snapshots = new List<ImageSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddImageSnapshots(snapshots, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                Dictionary<HeaderPart, string> headerPartKeys = CreateHeaderPartKeys(mainPart);
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (!headerPartKeys.TryGetValue(headerPart, out string? headerPartKey)) {
                        continue;
                    }

                    AddImageSnapshots(snapshots, headerPart, headerPart.Header, headerPartKey, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                Dictionary<FooterPart, string> footerPartKeys = CreateFooterPartKeys(mainPart);
                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (!footerPartKeys.TryGetValue(footerPart, out string? footerPartKey)) {
                        continue;
                    }

                    AddImageSnapshots(snapshots, footerPart, footerPart.Footer, footerPartKey, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                int footnoteIndex = 0;
                foreach (Footnote footnote in mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>()) {
                    if (!IsVisibleNote(footnote)) {
                        continue;
                    }

                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddImageSnapshots(snapshots, mainPart.FootnotesPart, footnote, FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                    footnoteIndex++;
                }

                int endnoteIndex = 0;
                foreach (Endnote endnote in mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? Enumerable.Empty<Endnote>()) {
                    if (!IsVisibleNote(endnote)) {
                        continue;
                    }

                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddImageSnapshots(snapshots, mainPart.EndnotesPart, endnote, EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                    endnoteIndex++;
                }
            }

            return snapshots;
        }

        private static void AddImageSnapshots(List<ImageSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase) {
            if (part == null || container == null) {
                return;
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                switch (ordered.Element) {
                    case DocumentFormat.OpenXml.Wordprocessing.Drawing drawing:
                        DocumentFormat.OpenXml.Drawing.Blip? blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                        if (blip == null) {
                            break;
                        }

                        string drawingVisualSignature = GetDrawingVisualSignature(part, drawing);
                        if (blip.Embed?.Value is string embeddedRelationshipId) {
                            AddEmbeddedImageSnapshot(snapshots, part, embeddedRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, GetImagePositionKey(partKey, ordered.DocumentOrder));
                        } else if (blip.Link?.Value is string externalRelationshipId) {
                            AddExternalImageSnapshot(snapshots, part, externalRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, GetImagePositionKey(partKey, ordered.DocumentOrder));
                        }

                        break;
                    case V.ImageData imageData when imageData.RelationshipId?.Value is string relationshipId:
                        string vmlVisualSignature = GetVmlVisualSignature(part, imageData);
                        if (part.ExternalRelationships.Any(item => item.Id == relationshipId)) {
                            AddExternalImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, GetImagePositionKey(partKey, ordered.DocumentOrder));
                        } else {
                            AddEmbeddedImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, GetImagePositionKey(partKey, ordered.DocumentOrder));
                        }

                        break;
                    }
                }
        }

        private static void AddEmbeddedImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            OpenXmlPart relatedPart;
            try {
                relatedPart = part.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return;
            }

            if (relatedPart is not ImagePart imagePart) {
                return;
            }

            using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            snapshots.Add(ImageSnapshot.FromEmbedded(CreateImageFingerprint(stream), visualSignature, partKey, documentOrder, positionKey));
        }

        private static void AddExternalImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            ExternalRelationship? relationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId);
            snapshots.Add(ImageSnapshot.FromExternal(relationship?.Uri?.ToString() ?? relationshipId, visualSignature, partKey, documentOrder, positionKey));
        }

        private static string GetDrawingVisualSignature(OpenXmlPart part, DocumentFormat.OpenXml.Wordprocessing.Drawing drawing) {
            OpenXmlElement clone = drawing.CloneNode(true);
            foreach (DocumentFormat.OpenXml.Drawing.Blip blip in clone.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                blip.Embed = null;
                blip.Link = null;
            }

            foreach (OpenXmlElement element in clone.Descendants()) {
                element.RemoveAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                element.RemoveAttribute("link", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            }

            foreach (DW.DocProperties properties in clone.Descendants<DW.DocProperties>()) {
                properties.Id = 0U;
                properties.Name = string.Empty;
            }

            foreach (PIC.NonVisualDrawingProperties properties in clone.Descendants<PIC.NonVisualDrawingProperties>()) {
                properties.Id = 0U;
                properties.Name = string.Empty;
            }

            return clone.OuterXml + GetImageHyperlinkSignature(part, drawing);
        }

        private static string GetVmlVisualSignature(OpenXmlPart part, V.ImageData imageData) {
            OpenXmlElement clone = (imageData.Parent ?? imageData).CloneNode(true);
            if (clone is V.ImageData clonedImageData) {
                clonedImageData.RelationshipId = null;
            }

            foreach (V.ImageData descendant in clone.Descendants<V.ImageData>()) {
                descendant.RelationshipId = null;
            }

            return clone.OuterXml + GetImageHyperlinkSignature(part, imageData);
        }

        private static string GetImageHyperlinkSignature(OpenXmlPart part, OpenXmlElement imageElement) {
            Hyperlink? hyperlink = imageElement.Ancestors<Hyperlink>().FirstOrDefault();
            return hyperlink == null ? string.Empty : "|hyperlink:" + GetHyperlinkSignature(part, hyperlink);
        }

        private static string GetImagePositionKey(string partKey, int documentOrder) {
            return partKey + ":" + documentOrder.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static IEnumerable<OrderedElement> EnumerateDescendantsWithOrder(OpenXmlElement? container, int orderBase) {
            if (container == null) {
                yield break;
            }

            int order = orderBase;
            foreach (OpenXmlElement element in container.Descendants()) {
                yield return new OrderedElement(element, order);
                order++;
            }
        }

        private static string ImageLocation(int imageIndex) {
            return "image[" + imageIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
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

        private static ImageFingerprint CreateImageFingerprint(Stream stream) {
            using System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create();
            byte[] buffer = new byte[81920];
            long length = 0;
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0) {
                sha256.TransformBlock(buffer, 0, bytesRead, null, 0);
                length += bytesRead;
            }

            sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return new ImageFingerprint(length, Convert.ToBase64String(sha256.Hash ?? Array.Empty<byte>()));
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

        private sealed class ImageSnapshot {
            private ImageSnapshot(ImageFingerprint? embeddedFingerprint, string? externalUri, string visualSignature, string partKey, int documentOrder, string positionKey) {
                EmbeddedFingerprint = embeddedFingerprint;
                ExternalUri = externalUri;
                VisualSignature = visualSignature;
                PartKey = partKey;
                DocumentOrder = documentOrder;
                PositionKey = positionKey;
            }

            internal ImageFingerprint? EmbeddedFingerprint { get; }

            internal string? ExternalUri { get; }

            internal string VisualSignature { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }

            internal string PositionKey { get; }

            internal string DisplayText => ExternalUri == null ? "[Image]" : "[Image: " + ExternalUri + "]";

            internal static ImageSnapshot FromEmbedded(ImageFingerprint embeddedFingerprint, string visualSignature, string partKey, int documentOrder, string positionKey) {
                return new ImageSnapshot(embeddedFingerprint, null, visualSignature, partKey, documentOrder, positionKey);
            }

            internal static ImageSnapshot FromExternal(string externalUri, string visualSignature, string partKey, int documentOrder, string positionKey) {
                return new ImageSnapshot(null, externalUri, visualSignature, partKey, documentOrder, positionKey);
            }
        }

        private sealed class ImageSnapshotEqualityComparer : IEqualityComparer<ImageSnapshot> {
            internal static readonly ImageSnapshotEqualityComparer Instance = new();

            public bool Equals(ImageSnapshot? x, ImageSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                if (x.ExternalUri != null || y.ExternalUri != null) {
                    return string.Equals(x.PartKey, y.PartKey, StringComparison.Ordinal) &&
                           string.Equals(x.ExternalUri, y.ExternalUri, StringComparison.Ordinal) &&
                           string.Equals(x.VisualSignature, y.VisualSignature, StringComparison.Ordinal);
                }

                return x.EmbeddedFingerprint != null &&
                       y.EmbeddedFingerprint != null &&
                       string.Equals(x.PartKey, y.PartKey, StringComparison.Ordinal) &&
                       x.EmbeddedFingerprint.Equals(y.EmbeddedFingerprint) &&
                       string.Equals(x.VisualSignature, y.VisualSignature, StringComparison.Ordinal);
            }

            public int GetHashCode(ImageSnapshot obj) {
                if (obj.ExternalUri != null) {
                    int externalHash = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    externalHash = (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.ExternalUri);
                    return (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }

                if (obj.EmbeddedFingerprint == null) {
                    return StringComparer.Ordinal.GetHashCode(obj.PartKey);
                }

                unchecked {
                    int hashCode = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    hashCode = (hashCode * 397) ^ obj.EmbeddedFingerprint.GetHashCode();
                    return (hashCode * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }
            }
        }

        private readonly struct ImageFingerprint : IEquatable<ImageFingerprint> {
            internal ImageFingerprint(long length, string sha256) {
                Length = length;
                Sha256 = sha256;
            }

            internal long Length { get; }

            internal string Sha256 { get; }

            public bool Equals(ImageFingerprint other) {
                return Length == other.Length &&
                       string.Equals(Sha256, other.Sha256, StringComparison.Ordinal);
            }

            public override bool Equals(object? obj) {
                return obj is ImageFingerprint other && Equals(other);
            }

            public override int GetHashCode() {
                unchecked {
                    return (Length.GetHashCode() * 397) ^ StringComparer.Ordinal.GetHashCode(Sha256);
                }
            }
        }
    }
}
