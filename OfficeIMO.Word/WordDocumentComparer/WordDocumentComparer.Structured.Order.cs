using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeBlockOrder(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            List<BlockOrderSnapshot> sourceBlocks = GetBlockOrderSnapshots(source, options);
            List<BlockOrderSnapshot> targetBlocks = GetBlockOrderSnapshots(target, options);
            if (sourceBlocks.Count < 2 || sourceBlocks.Count != targetBlocks.Count) {
                return;
            }

            string[] sourceKeys = sourceBlocks.Select(block => block.MatchKey).ToArray();
            string[] targetKeys = targetBlocks.Select(block => block.MatchKey).ToArray();
            if (sourceKeys.SequenceEqual(targetKeys, StringComparer.Ordinal) ||
                !HaveSameMultiset(sourceKeys, targetKeys)) {
                return;
            }

            result.Add(new WordComparisonFinding(
                WordComparisonScope.Paragraph,
                WordComparisonChangeKind.Modified,
                "document-order",
                null,
                null,
                string.Join(" -> ", sourceBlocks.Select(block => block.DisplayText).ToArray()),
                string.Join(" -> ", targetBlocks.Select(block => block.DisplayText).ToArray()),
                "Document block order changed."),
                targetBlocks[0].DocumentOrder);
        }

        private static List<BlockOrderSnapshot> GetBlockOrderSnapshots(WordDocument document, WordComparisonOptions options) {
            var snapshots = new List<BlockOrderSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddBlockOrderSnapshots(snapshots, document, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddBlockOrderSnapshots(snapshots, document, headerPartKey.Key, headerPartKey.Key.Header, headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddBlockOrderSnapshots(snapshots, document, footerPartKey.Key, footerPartKey.Key.Footer, footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, document, mainPart.FootnotesPart, footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, document, mainPart.EndnotesPart, endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddBlockOrderSnapshots(List<BlockOrderSnapshot> snapshots, WordDocument document, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                switch (ordered.Element) {
                    case Paragraph paragraph:
                        if (paragraph.Ancestors<TableCell>().Any()) {
                            break;
                        }

                        string paragraphText = GetParagraphText(paragraph);
                        if (paragraphText.Length == 0 && HasImageContent(paragraph)) {
                            break;
                        }

                        snapshots.Add(new BlockOrderSnapshot(
                            "paragraph:" + partKey + ":" + GetParagraphMatchText(paragraph, part, options),
                            "paragraph:" + paragraphText,
                            ordered.DocumentOrder));
                        break;
                    case Table table:
                        if (table.Ancestors<TableCell>().Any()) {
                            break;
                        }

                        var wordTable = new WordTable(document, table);
                        snapshots.Add(new BlockOrderSnapshot(
                            "table:" + GetTableMatchKey(partKey, wordTable, part, options),
                            "table:" + GetTableText(wordTable),
                            ordered.DocumentOrder));
                        break;
                    case TableCell cell:
                        AddTableCellBlockOrderSnapshots(snapshots, document, part, cell, partKey, ordered.DocumentOrder, options);
                        break;
                }
            }
        }

        private static void AddTableCellBlockOrderSnapshots(List<BlockOrderSnapshot> snapshots, WordDocument document, OpenXmlPart? part, TableCell cell, string partKey, int orderBase, WordComparisonOptions options) {
            int blockIndex = 0;
            string cellKey = partKey + ":cell:" + GetStableElementPath(cell);
            foreach (OpenXmlElement child in cell.Elements()) {
                switch (child) {
                    case Paragraph paragraph:
                        string paragraphText = GetParagraphText(paragraph);
                        if (paragraphText.Length == 0 && HasImageContent(paragraph)) {
                            break;
                        }

                        snapshots.Add(new BlockOrderSnapshot(
                            "cell-paragraph:" + cellKey + ":" + GetParagraphMatchText(paragraph, part, options),
                            "cell-paragraph:" + paragraphText,
                            orderBase + blockIndex));
                        break;
                    case Table table:
                        var wordTable = new WordTable(document, table);
                        snapshots.Add(new BlockOrderSnapshot(
                            "cell-table:" + cellKey + ":" + GetTableMatchKey(partKey, wordTable, part, options),
                            "cell-table:" + GetTableText(wordTable),
                            orderBase + blockIndex));
                        break;
                }

                blockIndex++;
            }
        }

        private static bool HaveSameMultiset(IReadOnlyList<string> source, IReadOnlyList<string> target) {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string key in source) {
                counts.TryGetValue(key, out int count);
                counts[key] = count + 1;
            }

            foreach (string key in target) {
                if (!counts.TryGetValue(key, out int count)) {
                    return false;
                }

                if (count == 1) {
                    counts.Remove(key);
                } else {
                    counts[key] = count - 1;
                }
            }

            return counts.Count == 0;
        }

        private readonly struct BlockOrderSnapshot {
            internal BlockOrderSnapshot(string matchKey, string displayText, int documentOrder) {
                MatchKey = matchKey;
                DisplayText = displayText;
                DocumentOrder = documentOrder;
            }

            internal string MatchKey { get; }

            internal string DisplayText { get; }

            internal int DocumentOrder { get; }
        }
    }
}
