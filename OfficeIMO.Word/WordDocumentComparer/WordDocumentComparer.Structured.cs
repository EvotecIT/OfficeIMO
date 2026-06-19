using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
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
            return result;
        }

        private static void AnalyzeParagraphs(WordDocument source, WordDocument target, WordComparisonResult result) {
            List<WordParagraph> sourceParagraphs = source.Paragraphs.Where(paragraph => !IsInsideTable(paragraph) && !string.IsNullOrEmpty(paragraph.Text)).ToList();
            List<WordParagraph> targetParagraphs = target.Paragraphs.Where(paragraph => !IsInsideTable(paragraph) && !string.IsNullOrEmpty(paragraph.Text)).ToList();
            IReadOnlyList<MatchedIndexPair> matchedParagraphs = FindMatchingIndexes(
                sourceParagraphs.Select(paragraph => paragraph.Text).ToList(),
                targetParagraphs.Select(paragraph => paragraph.Text).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedParagraphs) {
                AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, sourceParagraphs.Count, targetStart, targetParagraphs.Count, result);
        }

        private static void AnalyzeTables(WordDocument source, WordDocument target, WordComparisonResult result) {
            int tableCount = Math.Min(source.Tables.Count, target.Tables.Count);

            for (int tableIndex = 0; tableIndex < tableCount; tableIndex++) {
                AnalyzeTable(source.Tables[tableIndex], target.Tables[tableIndex], tableIndex, result);
            }

            for (int tableIndex = tableCount; tableIndex < target.Tables.Count; tableIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Table,
                    WordComparisonChangeKind.Inserted,
                    TableLocation(tableIndex),
                    null,
                    tableIndex,
                    null,
                    GetTableText(target.Tables[tableIndex]),
                    "Table inserted."));
            }

            for (int tableIndex = tableCount; tableIndex < source.Tables.Count; tableIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Table,
                    WordComparisonChangeKind.Deleted,
                    TableLocation(tableIndex),
                    tableIndex,
                    null,
                    GetTableText(source.Tables[tableIndex]),
                    null,
                    "Table deleted."));
            }
        }

        private static void AnalyzeTable(WordTable source, WordTable target, int tableIndex, WordComparisonResult result) {
            List<WordTableRow> sourceRows = source.Rows.ToList();
            List<WordTableRow> targetRows = target.Rows.ToList();
            IReadOnlyList<MatchedIndexPair> matchedRows = FindMatchingIndexes(
                sourceRows.Select(GetRowText).ToList(),
                targetRows.Select(GetRowText).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedRows) {
                AddTableRowRangeFindings(sourceRows, targetRows, tableIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRowRangeFindings(sourceRows, targetRows, tableIndex, sourceStart, sourceRows.Count, targetStart, targetRows.Count, result);
        }

        private static void AddParagraphRangeFindings(
            IReadOnlyList<WordParagraph> sourceParagraphs,
            IReadOnlyList<WordParagraph> targetParagraphs,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceCount = sourceEnd - sourceStart;
            int targetCount = targetEnd - targetStart;
            int pairedCount = Math.Min(sourceCount, targetCount);

            for (int offset = 0; offset < pairedCount; offset++) {
                int sourceIndex = sourceStart + offset;
                int targetIndex = targetStart + offset;
                string sourceText = sourceParagraphs[sourceIndex].Text;
                string targetText = targetParagraphs[targetIndex].Text;

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Modified,
                        ParagraphLocation(targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        "Paragraph text changed."));
                }
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                int targetIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Inserted,
                    ParagraphLocation(targetIndex),
                    null,
                    targetIndex,
                    null,
                    targetParagraphs[targetIndex].Text,
                    "Paragraph inserted."));
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                int sourceIndex = sourceStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Deleted,
                    ParagraphLocation(sourceIndex),
                    sourceIndex,
                    null,
                    sourceParagraphs[sourceIndex].Text,
                    null,
                    "Paragraph deleted."));
            }
        }

        private static void AddTableRowRangeFindings(
            IReadOnlyList<WordTableRow> sourceRows,
            IReadOnlyList<WordTableRow> targetRows,
            int tableIndex,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceCount = sourceEnd - sourceStart;
            int targetCount = targetEnd - targetStart;
            int pairedCount = Math.Min(sourceCount, targetCount);

            for (int offset = 0; offset < pairedCount; offset++) {
                int sourceIndex = sourceStart + offset;
                int targetIndex = targetStart + offset;
                AnalyzeTableRow(sourceRows[sourceIndex], targetRows[targetIndex], tableIndex, sourceIndex, targetIndex, result);
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                int rowIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableRow,
                    WordComparisonChangeKind.Inserted,
                    RowLocation(tableIndex, rowIndex),
                    null,
                    rowIndex,
                    null,
                    GetRowText(targetRows[rowIndex]),
                    "Table row inserted."));
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                int rowIndex = sourceStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableRow,
                    WordComparisonChangeKind.Deleted,
                    RowLocation(tableIndex, rowIndex),
                    rowIndex,
                    null,
                    GetRowText(sourceRows[rowIndex]),
                    null,
                    "Table row deleted."));
            }
        }

        private static void AnalyzeTableRow(WordTableRow source, WordTableRow target, int tableIndex, int sourceRowIndex, int targetRowIndex, WordComparisonResult result) {
            int cellCount = Math.Min(source.CellsCount, target.CellsCount);

            for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
                string sourceText = GetCellText(source.Cells[cellIndex]);
                string targetText = GetCellText(target.Cells[cellIndex]);

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.TableCell,
                        WordComparisonChangeKind.Modified,
                        CellLocation(tableIndex, targetRowIndex, cellIndex),
                        cellIndex,
                        cellIndex,
                        sourceText,
                        targetText,
                        "Table cell text changed."));
                }
            }

            for (int cellIndex = cellCount; cellIndex < target.CellsCount; cellIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableCell,
                    WordComparisonChangeKind.Inserted,
                    CellLocation(tableIndex, targetRowIndex, cellIndex),
                    null,
                    cellIndex,
                    null,
                    GetCellText(target.Cells[cellIndex]),
                    "Table cell inserted."));
            }

            for (int cellIndex = cellCount; cellIndex < source.CellsCount; cellIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableCell,
                    WordComparisonChangeKind.Deleted,
                    CellLocation(tableIndex, sourceRowIndex, cellIndex),
                    cellIndex,
                    null,
                    GetCellText(source.Cells[cellIndex]),
                    null,
                    "Table cell deleted."));
            }
        }

        private static void AnalyzeImages(WordDocument source, WordDocument target, WordComparisonResult result) {
            IReadOnlyList<byte[]> sourceImages = source.GetImages();
            IReadOnlyList<byte[]> targetImages = target.GetImages();
            IReadOnlyList<MatchedIndexPair> matchedImages = FindMatchingIndexes(
                sourceImages,
                targetImages,
                ByteArrayEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedImages) {
                AddImageRangeFindings(sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddImageRangeFindings(sourceStart, sourceImages.Count, targetStart, targetImages.Count, result);
        }

        private static void AddImageRangeFindings(int sourceStart, int sourceEnd, int targetStart, int targetEnd, WordComparisonResult result) {
            int sourceCount = sourceEnd - sourceStart;
            int targetCount = targetEnd - targetStart;
            int pairedCount = Math.Min(sourceCount, targetCount);

            for (int offset = 0; offset < pairedCount; offset++) {
                int sourceIndex = sourceStart + offset;
                int targetIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Modified,
                    ImageLocation(targetIndex),
                    sourceIndex,
                    targetIndex,
                    "[Image]",
                    "[Image]",
                    "Image payload changed."));
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                int imageIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Inserted,
                    ImageLocation(imageIndex),
                    null,
                    imageIndex,
                    null,
                    "[Image]",
                    "Image inserted."));
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                int imageIndex = sourceStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Deleted,
                    ImageLocation(imageIndex),
                    imageIndex,
                    null,
                    "[Image]",
                    null,
                    "Image deleted."));
            }
        }

        private static IReadOnlyList<MatchedIndexPair> FindMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) {
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

        private static bool IsInsideTable(WordParagraph paragraph) {
            return paragraph._paragraph.Ancestors<TableCell>().Any();
        }

        private static string GetTableText(WordTable table) {
            return string.Join(Environment.NewLine, table.Rows.Select(GetRowText).Where(text => !string.IsNullOrEmpty(text)).ToArray());
        }

        private static string GetRowText(WordTableRow row) {
            return string.Join(" | ", row.Cells.Select(GetCellText).ToArray());
        }

        private static string GetCellText(WordTableCell cell) {
            return string.Join(Environment.NewLine, cell.Paragraphs.Select(paragraph => paragraph.Text).Where(text => !string.IsNullOrEmpty(text)).ToArray());
        }

        private static string ParagraphLocation(int paragraphIndex) {
            return "paragraph[" + paragraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
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

        private readonly struct MatchedIndexPair {
            internal MatchedIndexPair(int sourceIndex, int targetIndex) {
                SourceIndex = sourceIndex;
                TargetIndex = targetIndex;
            }

            internal int SourceIndex { get; }

            internal int TargetIndex { get; }
        }

        private sealed class ByteArrayEqualityComparer : IEqualityComparer<byte[]> {
            internal static readonly ByteArrayEqualityComparer Instance = new();

            public bool Equals(byte[]? x, byte[]? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return x.SequenceEqual(y);
            }

            public int GetHashCode(byte[] obj) {
                if (obj == null) {
                    return 0;
                }

                unchecked {
                    int hashCode = 17;
                    foreach (byte value in obj) {
                        hashCode = (hashCode * 31) + value;
                    }

                    return hashCode;
                }
            }
        }
    }
}
