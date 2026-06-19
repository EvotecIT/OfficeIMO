using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeBlockOrder(WordDocument source, WordDocument target, WordComparisonResult result) {
            List<BlockOrderSnapshot> sourceBlocks = GetBlockOrderSnapshots(source);
            List<BlockOrderSnapshot> targetBlocks = GetBlockOrderSnapshots(target);
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

        private static List<BlockOrderSnapshot> GetBlockOrderSnapshots(WordDocument document) {
            var snapshots = new List<BlockOrderSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddBlockOrderSnapshots(snapshots, document, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                Dictionary<HeaderPart, string> headerPartKeys = CreateHeaderPartKeys(mainPart);
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (!headerPartKeys.TryGetValue(headerPart, out string? headerPartKey)) {
                        continue;
                    }

                    AddBlockOrderSnapshots(snapshots, document, headerPart, headerPart.Header, headerPartKey, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                Dictionary<FooterPart, string> footerPartKeys = CreateFooterPartKeys(mainPart);
                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (!footerPartKeys.TryGetValue(footerPart, out string? footerPartKey)) {
                        continue;
                    }

                    AddBlockOrderSnapshots(snapshots, document, footerPart, footerPart.Footer, footerPartKey, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                int footnoteIndex = 0;
                foreach (Footnote footnote in mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>()) {
                    if (!IsVisibleNote(footnote)) {
                        continue;
                    }

                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, document, mainPart.FootnotesPart, footnote, FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                    footnoteIndex++;
                }

                int endnoteIndex = 0;
                foreach (Endnote endnote in mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? Enumerable.Empty<Endnote>()) {
                    if (!IsVisibleNote(endnote)) {
                        continue;
                    }

                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, document, mainPart.EndnotesPart, endnote, EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                    endnoteIndex++;
                }
            }

            return snapshots;
        }

        private static void AddBlockOrderSnapshots(List<BlockOrderSnapshot> snapshots, WordDocument document, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase) {
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
                            "paragraph:" + partKey + ":" + GetParagraphMatchText(paragraph, part),
                            "paragraph:" + paragraphText,
                            ordered.DocumentOrder));
                        break;
                    case Table table:
                        if (table.Ancestors<TableCell>().Any()) {
                            break;
                        }

                        var wordTable = new WordTable(document, table);
                        snapshots.Add(new BlockOrderSnapshot(
                            "table:" + GetTableMatchKey(partKey, wordTable, part),
                            "table:" + GetTableText(wordTable),
                            ordered.DocumentOrder));
                        break;
                }
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
