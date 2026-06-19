using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const int LcsCellLimit = 1_000_000;
        private const int BodyPartOrderBase = 0;
        private const int HeaderPartOrderBase = 1_000_000;
        private const int FooterPartOrderBase = 2_000_000;
        private const int RelatedPartOrderStride = 100_000;
        private const string BodyPartKey = "body";
        private const string HeaderPartKeyPrefix = "header:";
        private const string FooterPartKeyPrefix = "footer:";
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
            int sourceCount = sourceEnd - sourceStart;
            int targetCount = targetEnd - targetStart;
            int pairedCount = Math.Min(sourceCount, targetCount);

            for (int offset = 0; offset < pairedCount; offset++) {
                int sourceIndex = sourceStart + offset;
                int targetIndex = targetStart + offset;
                if (!string.Equals(sourceTables[sourceIndex].PartKey, targetTables[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedTableFinding(sourceTables, sourceIndex, result);
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    continue;
                }

                AnalyzeTable(sourceTables[sourceIndex].Table, targetTables[targetIndex].Table, targetIndex, targetTables[targetIndex].DocumentOrder, result);
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                AddInsertedTableFinding(targetTables, targetStart + offset, result);
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                AddDeletedTableFinding(sourceTables, sourceStart + offset, result);
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

        private static void AnalyzeTable(WordTable source, WordTable target, int tableIndex, int tableDocumentOrder, WordComparisonResult result) {
            List<WordTableRow> sourceRows = source.Rows.ToList();
            List<WordTableRow> targetRows = target.Rows.ToList();
            IReadOnlyList<MatchedIndexPair> matchedRows = FindMatchingIndexes(
                sourceRows.Select(GetRowMatchKey).ToList(),
                targetRows.Select(GetRowMatchKey).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedRows) {
                AddTableRowRangeFindings(sourceRows, targetRows, tableIndex, tableDocumentOrder, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRowRangeFindings(sourceRows, targetRows, tableIndex, tableDocumentOrder, sourceStart, sourceRows.Count, targetStart, targetRows.Count, result);
        }

        private static void AddTableRowRangeFindings(
            IReadOnlyList<WordTableRow> sourceRows,
            IReadOnlyList<WordTableRow> targetRows,
            int tableIndex,
            int tableDocumentOrder,
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
                if (sourceRows[sourceIndex].Cells.Count != targetRows[targetIndex].Cells.Count &&
                    string.Equals(GetRowText(sourceRows[sourceIndex]), GetRowText(targetRows[targetIndex]), StringComparison.Ordinal)) {
                    AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                    AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                    continue;
                }

                AnalyzeTableRow(sourceRows[sourceIndex], targetRows[targetIndex], tableIndex, tableDocumentOrder, sourceIndex, targetIndex, result);
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetStart + offset, result);
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceStart + offset, result);
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
                tableDocumentOrder + rowIndex);
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
                tableDocumentOrder + rowIndex);
        }

        private static void AnalyzeTableRow(WordTableRow source, WordTableRow target, int tableIndex, int tableDocumentOrder, int sourceRowIndex, int targetRowIndex, WordComparisonResult result) {
            List<WordTableCell> sourceCells = source.Cells.ToList();
            List<WordTableCell> targetCells = target.Cells.ToList();
            IReadOnlyList<MatchedIndexPair> matchedCells = FindMatchingIndexes(
                sourceCells.Select(GetCellText).ToList(),
                targetCells.Select(GetCellText).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedCells) {
                AddTableCellRangeFindings(sourceCells, targetCells, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableCellRangeFindings(sourceCells, targetCells, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, sourceCells.Count, targetStart, targetCells.Count, result);
        }

        private static void AddTableCellRangeFindings(
            IReadOnlyList<WordTableCell> sourceCells,
            IReadOnlyList<WordTableCell> targetCells,
            int tableIndex,
            int tableDocumentOrder,
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
                        "Table cell text changed."),
                        RowLocationOrder(tableDocumentOrder, targetRowIndex, targetIndex));
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
                    "Table cell inserted."),
                    RowLocationOrder(tableDocumentOrder, targetRowIndex, cellIndex));
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
                    "Table cell deleted."),
                    RowLocationOrder(tableDocumentOrder, sourceRowIndex, cellIndex));
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
                if (!string.Equals(sourceImages[sourceIndex].PartKey, targetImages[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedImageFinding(sourceImages, sourceIndex, result);
                    AddInsertedImageFinding(targetImages, targetIndex, result);
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
            }

            for (int offset = pairedCount; offset < targetCount; offset++) {
                AddInsertedImageFinding(targetImages, targetStart + offset, result);
            }

            for (int offset = pairedCount; offset < sourceCount; offset++) {
                AddDeletedImageFinding(sourceImages, sourceStart + offset, result);
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
            AddTableSnapshots(snapshots, document, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddTableSnapshots(snapshots, document, headerPart.Header, HeaderPartKeyPrefix + headerIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddTableSnapshots(snapshots, document, footerPart.Footer, FooterPartKeyPrefix + footerIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }
            }

            return snapshots;
        }

        private static void AddTableSnapshots(List<TableSnapshot> snapshots, WordDocument document, OpenXmlElement? container, string partKey, int orderBase) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Table table) {
                    continue;
                }

                var wordTable = new WordTable(document, table);
                string text = GetTableText(wordTable);
                snapshots.Add(new TableSnapshot(wordTable, text, GetTableMatchKey(partKey, wordTable, text), partKey, ordered.DocumentOrder));
            }
        }

        private static string GetTableText(WordTable table) {
            return string.Join(TableRowSeparator, table.Rows.Select(GetRowText).ToArray());
        }

        private static string GetRowText(WordTableRow row) {
            return string.Join(" | ", row.Cells.Select(GetCellText).ToArray());
        }

        private static string GetTableMatchKey(string partKey, WordTable table, string text) {
            return partKey + TableRowSeparator + GetTableShape(table) + TableRowSeparator + text;
        }

        private static string GetTableShape(WordTable table) {
            return string.Join(";", table.Rows.Select(GetRowShape).ToArray());
        }

        private static string GetRowMatchKey(WordTableRow row) {
            return GetRowShape(row) + TableRowSeparator + GetRowText(row);
        }

        private static string GetRowShape(WordTableRow row) {
            return row.Cells.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetCellText(WordTableCell cell) {
            return string.Join(
                CellParagraphSeparator,
                cell._tableCell.Descendants<Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<TableCell>().FirstOrDefault(), cell._tableCell))
                    .Select(GetParagraphText)
                    .ToArray());
        }

        private static List<ImageSnapshot> GetImageSnapshots(WordDocument document) {
            var snapshots = new List<ImageSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddImageSnapshots(snapshots, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddImageSnapshots(snapshots, headerPart, headerPart.Header, HeaderPartKeyPrefix + headerIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddImageSnapshots(snapshots, footerPart, footerPart.Footer, FooterPartKeyPrefix + footerIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }
            }

            return snapshots;
        }

        private static void AddImageSnapshots(List<ImageSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase) {
            if (part == null || container == null) {
                return;
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not DocumentFormat.OpenXml.Wordprocessing.Drawing drawing) {
                    continue;
                }

                DocumentFormat.OpenXml.Drawing.Blip? blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                if (blip == null) {
                    continue;
                }

                string visualSignature = GetDrawingVisualSignature(drawing);
                if (blip.Embed?.Value is string embeddedRelationshipId) {
                    AddEmbeddedImageSnapshot(snapshots, part, embeddedRelationshipId, visualSignature, partKey, ordered.DocumentOrder);
                    continue;
                }

                if (blip.Link?.Value is string externalRelationshipId) {
                    AddExternalImageSnapshot(snapshots, part, externalRelationshipId, visualSignature, partKey, ordered.DocumentOrder);
                }
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not V.ImageData imageData) {
                    continue;
                }

                if (imageData.RelationshipId?.Value is string relationshipId) {
                    string visualSignature = GetVmlVisualSignature(imageData);
                    if (part.ExternalRelationships.Any(item => item.Id == relationshipId)) {
                        AddExternalImageSnapshot(snapshots, part, relationshipId, visualSignature, partKey, ordered.DocumentOrder);
                    } else {
                        AddEmbeddedImageSnapshot(snapshots, part, relationshipId, visualSignature, partKey, ordered.DocumentOrder);
                    }
                }
            }
        }

        private static void AddEmbeddedImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder) {
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
            using var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            snapshots.Add(ImageSnapshot.FromEmbedded(memoryStream.ToArray(), visualSignature, partKey, documentOrder));
        }

        private static void AddExternalImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder) {
            ExternalRelationship? relationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId);
            snapshots.Add(ImageSnapshot.FromExternal(relationship?.Uri?.ToString() ?? relationshipId, visualSignature, partKey, documentOrder));
        }

        private static string GetDrawingVisualSignature(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing) {
            OpenXmlElement clone = drawing.CloneNode(true);
            foreach (DocumentFormat.OpenXml.Drawing.Blip blip in clone.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                blip.Embed = null;
                blip.Link = null;
            }

            foreach (DW.DocProperties properties in clone.Descendants<DW.DocProperties>()) {
                properties.Id = 0U;
                properties.Name = string.Empty;
            }

            return clone.OuterXml;
        }

        private static string GetVmlVisualSignature(V.ImageData imageData) {
            OpenXmlElement clone = (imageData.Parent ?? imageData).CloneNode(true);
            if (clone is V.ImageData clonedImageData) {
                clonedImageData.RelationshipId = null;
            }

            foreach (V.ImageData descendant in clone.Descendants<V.ImageData>()) {
                descendant.RelationshipId = null;
            }

            return clone.OuterXml;
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

        private static int RowLocationOrder(int tableDocumentOrder, int rowIndex, int childIndex) {
            return tableDocumentOrder + rowIndex + childIndex;
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
            internal TableSnapshot(WordTable table, string text, string matchKey, string partKey, int documentOrder) {
                Table = table;
                Text = text;
                MatchKey = matchKey;
                PartKey = partKey;
                DocumentOrder = documentOrder;
            }

            internal WordTable Table { get; }

            internal string Text { get; }

            internal string MatchKey { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }
        }

        private sealed class ImageSnapshot {
            private ImageSnapshot(byte[]? embeddedBytes, string? externalUri, string visualSignature, string partKey, int documentOrder) {
                EmbeddedBytes = embeddedBytes;
                ExternalUri = externalUri;
                VisualSignature = visualSignature;
                PartKey = partKey;
                DocumentOrder = documentOrder;
            }

            internal byte[]? EmbeddedBytes { get; }

            internal string? ExternalUri { get; }

            internal string VisualSignature { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }

            internal string DisplayText => ExternalUri == null ? "[Image]" : "[Image: " + ExternalUri + "]";

            internal static ImageSnapshot FromEmbedded(byte[] embeddedBytes, string visualSignature, string partKey, int documentOrder) {
                return new ImageSnapshot(embeddedBytes, null, visualSignature, partKey, documentOrder);
            }

            internal static ImageSnapshot FromExternal(string externalUri, string visualSignature, string partKey, int documentOrder) {
                return new ImageSnapshot(null, externalUri, visualSignature, partKey, documentOrder);
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

                return x.EmbeddedBytes != null &&
                       y.EmbeddedBytes != null &&
                       string.Equals(x.PartKey, y.PartKey, StringComparison.Ordinal) &&
                       x.EmbeddedBytes.SequenceEqual(y.EmbeddedBytes) &&
                       string.Equals(x.VisualSignature, y.VisualSignature, StringComparison.Ordinal);
            }

            public int GetHashCode(ImageSnapshot obj) {
                if (obj.ExternalUri != null) {
                    int externalHash = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    externalHash = (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.ExternalUri);
                    return (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }

                if (obj.EmbeddedBytes == null) {
                    return StringComparer.Ordinal.GetHashCode(obj.PartKey);
                }

                unchecked {
                    int hashCode = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    foreach (byte value in obj.EmbeddedBytes) {
                        hashCode = (hashCode * 31) + value;
                    }

                    return (hashCode * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }
            }
        }
    }
}
