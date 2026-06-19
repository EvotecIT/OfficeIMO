using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

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
            List<ParagraphSnapshot> sourceParagraphs = GetLogicalBodyParagraphs(source);
            List<ParagraphSnapshot> targetParagraphs = GetLogicalBodyParagraphs(target);
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
            List<TableSnapshot> sourceTables = GetTableSnapshots(source);
            List<TableSnapshot> targetTables = GetTableSnapshots(target);
            IReadOnlyList<MatchedIndexPair> matchedTables = FindMatchingIndexes(
                sourceTables.Select(table => table.Text).ToList(),
                targetTables.Select(table => table.Text).ToList(),
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
            int sourceCount = sourceEnd - sourceStart;
            int targetCount = targetEnd - targetStart;
            int pairedCount = Math.Min(sourceCount, targetCount);

            for (int offset = 0; offset < pairedCount; offset++) {
                int sourceIndex = sourceStart + offset;
                int targetIndex = targetStart + offset;
                AnalyzeTable(sourceTables[sourceIndex].Table, targetTables[targetIndex].Table, targetIndex, result);
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                int tableIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Table,
                    WordComparisonChangeKind.Inserted,
                    TableLocation(tableIndex),
                    null,
                    tableIndex,
                    null,
                    targetTables[tableIndex].Text,
                    "Table inserted."));
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                int tableIndex = sourceStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Table,
                    WordComparisonChangeKind.Deleted,
                    TableLocation(tableIndex),
                    tableIndex,
                    null,
                    sourceTables[tableIndex].Text,
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
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
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
            List<WordTableCell> sourceCells = source.Cells.ToList();
            List<WordTableCell> targetCells = target.Cells.ToList();
            IReadOnlyList<MatchedIndexPair> matchedCells = FindMatchingIndexes(
                sourceCells.Select(GetCellText).ToList(),
                targetCells.Select(GetCellText).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedCells) {
                AddTableCellRangeFindings(sourceCells, targetCells, tableIndex, sourceRowIndex, targetRowIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableCellRangeFindings(sourceCells, targetCells, tableIndex, sourceRowIndex, targetRowIndex, sourceStart, sourceCells.Count, targetStart, targetCells.Count, result);
        }

        private static void AddTableCellRangeFindings(
            IReadOnlyList<WordTableCell> sourceCells,
            IReadOnlyList<WordTableCell> targetCells,
            int tableIndex,
            int sourceRowIndex,
            int targetRowIndex,
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
                string sourceText = GetCellText(sourceCells[sourceIndex]);
                string targetText = GetCellText(targetCells[targetIndex]);

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.TableCell,
                        WordComparisonChangeKind.Modified,
                        CellLocation(tableIndex, targetRowIndex, targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        "Table cell text changed."));
                }
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                int cellIndex = targetStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableCell,
                    WordComparisonChangeKind.Inserted,
                    CellLocation(tableIndex, targetRowIndex, cellIndex),
                    null,
                    cellIndex,
                    null,
                    GetCellText(targetCells[cellIndex]),
                    "Table cell inserted."));
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                int cellIndex = sourceStart + offset;
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.TableCell,
                    WordComparisonChangeKind.Deleted,
                    CellLocation(tableIndex, sourceRowIndex, cellIndex),
                    cellIndex,
                    null,
                    GetCellText(sourceCells[cellIndex]),
                    null,
                    "Table cell deleted."));
            }
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
        }

        private static void AddImageRangeFindings(
            IReadOnlyList<ImageSnapshot> sourceImages,
            IReadOnlyList<ImageSnapshot> targetImages,
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
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Modified,
                    ImageLocation(targetIndex),
                    sourceIndex,
                    targetIndex,
                    sourceImages[sourceIndex].DisplayText,
                    targetImages[targetIndex].DisplayText,
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
                    targetImages[imageIndex].DisplayText,
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
                    sourceImages[imageIndex].DisplayText,
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

        private static List<ParagraphSnapshot> GetLogicalBodyParagraphs(WordDocument document) {
            var snapshots = new List<ParagraphSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddParagraphSnapshots(snapshots, mainPart?.Document?.Body);

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddParagraphSnapshots(snapshots, headerPart.Header);
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddParagraphSnapshots(snapshots, footerPart.Footer);
                }
            }

            return snapshots;
        }

        private static void AddParagraphSnapshots(List<ParagraphSnapshot> snapshots, OpenXmlElement? container) {
            IEnumerable<Paragraph> paragraphs = container?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>();
            foreach (Paragraph paragraph in paragraphs) {
                if (paragraph.Ancestors<TableCell>().Any()) {
                    continue;
                }

                string text = GetParagraphText(paragraph);
                if (text.Length == 0 && paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any()) {
                    continue;
                }

                snapshots.Add(new ParagraphSnapshot(text));
            }
        }

        private static string GetParagraphText(Paragraph paragraph) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement element in paragraph.Descendants()) {
                switch (element) {
                    case Text text:
                        builder.Append(text.Text);
                        break;
                    case TabChar:
                        builder.Append('\t');
                        break;
                    case Break breakNode when breakNode.Type == null || breakNode.Type.Value == BreakValues.TextWrapping:
                        builder.Append('\n');
                        break;
                }
            }

            return builder.ToString();
        }

        private static List<TableSnapshot> GetTableSnapshots(WordDocument document) {
            var snapshots = new List<TableSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddTableSnapshots(snapshots, document, mainPart?.Document?.Body);

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddTableSnapshots(snapshots, document, headerPart.Header);
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddTableSnapshots(snapshots, document, footerPart.Footer);
                }
            }

            return snapshots;
        }

        private static void AddTableSnapshots(List<TableSnapshot> snapshots, WordDocument document, OpenXmlElement? container) {
            IEnumerable<Table> tables = container?.Descendants<Table>() ?? Enumerable.Empty<Table>();
            foreach (Table table in tables) {
                var wordTable = new WordTable(document, table);
                snapshots.Add(new TableSnapshot(wordTable, GetTableText(wordTable)));
            }
        }

        private static string GetTableText(WordTable table) {
            return string.Join(Environment.NewLine, table.Rows.Select(GetRowText).ToArray());
        }

        private static string GetRowText(WordTableRow row) {
            return string.Join(" | ", row.Cells.Select(GetCellText).ToArray());
        }

        private static string GetCellText(WordTableCell cell) {
            return string.Join(Environment.NewLine, cell._tableCell.ChildElements.OfType<Paragraph>().Select(GetParagraphText).ToArray());
        }

        private static List<ImageSnapshot> GetImageSnapshots(WordDocument document) {
            var snapshots = new List<ImageSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddImageSnapshots(snapshots, mainPart, mainPart?.Document?.Body);

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddImageSnapshots(snapshots, headerPart, headerPart.Header);
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddImageSnapshots(snapshots, footerPart, footerPart.Footer);
                }
            }

            return snapshots;
        }

        private static void AddImageSnapshots(List<ImageSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container) {
            if (part == null || container == null) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Wordprocessing.Drawing drawing in container.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>()) {
                DocumentFormat.OpenXml.Drawing.Blip? blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                if (blip == null) {
                    continue;
                }

                if (blip.Embed?.Value is string embeddedRelationshipId) {
                    if (part.GetPartById(embeddedRelationshipId) is ImagePart imagePart) {
                        using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                        using var memoryStream = new MemoryStream();
                        stream.CopyTo(memoryStream);
                        snapshots.Add(ImageSnapshot.FromEmbedded(memoryStream.ToArray()));
                    }
                    continue;
                }

                if (blip.Link?.Value is string externalRelationshipId) {
                    ExternalRelationship? relationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == externalRelationshipId);
                    snapshots.Add(ImageSnapshot.FromExternal(relationship?.Uri?.ToString() ?? externalRelationshipId));
                }
            }
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

        private sealed class ParagraphSnapshot {
            internal ParagraphSnapshot(string text) {
                Text = text;
            }

            internal string Text { get; }
        }

        private sealed class TableSnapshot {
            internal TableSnapshot(WordTable table, string text) {
                Table = table;
                Text = text;
            }

            internal WordTable Table { get; }

            internal string Text { get; }
        }

        private sealed class ImageSnapshot {
            private ImageSnapshot(byte[]? embeddedBytes, string? externalUri) {
                EmbeddedBytes = embeddedBytes;
                ExternalUri = externalUri;
            }

            internal byte[]? EmbeddedBytes { get; }

            internal string? ExternalUri { get; }

            internal string DisplayText => ExternalUri == null ? "[Image]" : "[Image: " + ExternalUri + "]";

            internal static ImageSnapshot FromEmbedded(byte[] embeddedBytes) {
                return new ImageSnapshot(embeddedBytes, null);
            }

            internal static ImageSnapshot FromExternal(string externalUri) {
                return new ImageSnapshot(null, externalUri);
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
                    return string.Equals(x.ExternalUri, y.ExternalUri, StringComparison.Ordinal);
                }

                return x.EmbeddedBytes != null && y.EmbeddedBytes != null && x.EmbeddedBytes.SequenceEqual(y.EmbeddedBytes);
            }

            public int GetHashCode(ImageSnapshot obj) {
                if (obj.ExternalUri != null) {
                    return StringComparer.Ordinal.GetHashCode(obj.ExternalUri);
                }

                if (obj.EmbeddedBytes == null) {
                    return 0;
                }

                unchecked {
                    int hashCode = 17;
                    foreach (byte value in obj.EmbeddedBytes) {
                        hashCode = (hashCode * 31) + value;
                    }

                    return hashCode;
                }
            }
        }
    }
}
