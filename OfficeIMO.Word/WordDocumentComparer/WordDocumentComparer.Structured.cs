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
            List<WordParagraph> sourceParagraphs = source.Paragraphs.Where(paragraph => !IsInsideTable(paragraph)).ToList();
            List<WordParagraph> targetParagraphs = target.Paragraphs.Where(paragraph => !IsInsideTable(paragraph)).ToList();
            int count = Math.Min(sourceParagraphs.Count, targetParagraphs.Count);

            for (int i = 0; i < count; i++) {
                string sourceText = sourceParagraphs[i].Text;
                string targetText = targetParagraphs[i].Text;

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Modified,
                        "paragraph[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                        i,
                        i,
                        sourceText,
                        targetText,
                        "Paragraph text changed."));
                }
            }

            for (int i = count; i < targetParagraphs.Count; i++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Inserted,
                    "paragraph[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                    null,
                    i,
                    null,
                    targetParagraphs[i].Text,
                    "Paragraph inserted."));
            }

            for (int i = count; i < sourceParagraphs.Count; i++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Deleted,
                    "paragraph[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                    i,
                    null,
                    sourceParagraphs[i].Text,
                    null,
                    "Paragraph deleted."));
            }
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
            int rowCount = Math.Min(source.RowsCount, target.RowsCount);

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                AnalyzeTableRow(source.Rows[rowIndex], target.Rows[rowIndex], tableIndex, rowIndex, result);
            }

            for (int rowIndex = rowCount; rowIndex < target.RowsCount; rowIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableRow,
                    WordComparisonChangeKind.Inserted,
                    RowLocation(tableIndex, rowIndex),
                    null,
                    rowIndex,
                    null,
                    GetRowText(target.Rows[rowIndex]),
                    "Table row inserted."));
            }

            for (int rowIndex = rowCount; rowIndex < source.RowsCount; rowIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableRow,
                    WordComparisonChangeKind.Deleted,
                    RowLocation(tableIndex, rowIndex),
                    rowIndex,
                    null,
                    GetRowText(source.Rows[rowIndex]),
                    null,
                    "Table row deleted."));
            }
        }

        private static void AnalyzeTableRow(WordTableRow source, WordTableRow target, int tableIndex, int rowIndex, WordComparisonResult result) {
            int cellCount = Math.Min(source.CellsCount, target.CellsCount);

            for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
                string sourceText = GetCellText(source.Cells[cellIndex]);
                string targetText = GetCellText(target.Cells[cellIndex]);

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.TableCell,
                        WordComparisonChangeKind.Modified,
                        CellLocation(tableIndex, rowIndex, cellIndex),
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
                    CellLocation(tableIndex, rowIndex, cellIndex),
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
                    CellLocation(tableIndex, rowIndex, cellIndex),
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
            int count = Math.Min(sourceImages.Count, targetImages.Count);

            for (int imageIndex = 0; imageIndex < count; imageIndex++) {
                if (!sourceImages[imageIndex].SequenceEqual(targetImages[imageIndex])) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Image,
                        WordComparisonChangeKind.Modified,
                        "image[" + imageIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                        imageIndex,
                        imageIndex,
                        "[Image]",
                        "[Image]",
                        "Image payload changed."));
                }
            }

            for (int imageIndex = count; imageIndex < targetImages.Count; imageIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Inserted,
                    "image[" + imageIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                    null,
                    imageIndex,
                    null,
                    "[Image]",
                    "Image inserted."));
            }

            for (int imageIndex = count; imageIndex < sourceImages.Count; imageIndex++) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Deleted,
                    "image[" + imageIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]",
                    imageIndex,
                    null,
                    "[Image]",
                    null,
                    "Image deleted."));
            }
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

        private static string TableLocation(int tableIndex) {
            return "table[" + tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string RowLocation(int tableIndex, int rowIndex) {
            return TableLocation(tableIndex) + "/row[" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string CellLocation(int tableIndex, int rowIndex, int cellIndex) {
            return RowLocation(tableIndex, rowIndex) + "/cell[" + cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }
    }
}
