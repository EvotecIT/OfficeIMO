using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeBlockOrder(WordPackageView source, WordPackageView target, WordComparisonResult result, WordComparisonOptions options) {
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

        private static List<BlockOrderSnapshot> GetBlockOrderSnapshots(WordPackageView document, WordComparisonOptions options) {
            var snapshots = new List<BlockOrderSnapshot>();
            WordPartView? mainPart = document.Main;
            AddBlockOrderSnapshots(snapshots, mainPart, BodyOf(mainPart), BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddBlockOrderSnapshots(snapshots, headerPartKey.Key, HeaderRoot(headerPartKey.Key), headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddBlockOrderSnapshots(snapshots, footerPartKey.Key, FooterRoot(footerPartKey.Key), footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, FootnotesPart(mainPart), footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddBlockOrderSnapshots(snapshots, EndnotesPart(mainPart), endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddBlockOrderSnapshots(List<BlockOrderSnapshot> snapshots, WordPartView? part, OfficeXmlNode container, string partKey, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                OfficeXmlNode element = ordered.Element;
                if (element.NamespaceId != W) {
                    continue;
                }

                switch (element.LocalName) {
                    case "p":
                        if (HasAncestorW(element, "tc")) {
                            break;
                        }

                        string paragraphText = GetParagraphText(element);
                        if (paragraphText.Length == 0 && HasImageContent(element)) {
                            break;
                        }

                        snapshots.Add(new BlockOrderSnapshot(
                            "paragraph:" + partKey + ":" + GetParagraphMatchText(element, part, options),
                            "paragraph:" + paragraphText,
                            ordered.DocumentOrder));
                        break;
                    case "tbl":
                        if (HasAncestorW(element, "tc")) {
                            break;
                        }

                        snapshots.Add(new BlockOrderSnapshot(
                            "table:" + GetTableMatchKey(partKey, element, part, options),
                            "table:" + GetTableText(element),
                            ordered.DocumentOrder));
                        break;
                    case "tc":
                        AddTableCellBlockOrderSnapshots(snapshots, part, element, partKey, ordered.DocumentOrder, options);
                        break;
                }
            }
        }

        private static void AddTableCellBlockOrderSnapshots(List<BlockOrderSnapshot> snapshots, WordPartView? part, OfficeXmlNode cell, string partKey, int orderBase, WordComparisonOptions options) {
            int blockIndex = 0;
            string cellKey = partKey + ":cell:" + GetStableElementPath(cell);
            for (OfficeXmlNode child = cell.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(W, "p")) {
                    string paragraphText = GetParagraphText(child);
                    if (!(paragraphText.Length == 0 && HasImageContent(child))) {
                        snapshots.Add(new BlockOrderSnapshot(
                            "cell-paragraph:" + cellKey + ":" + GetParagraphMatchText(child, part, options),
                            "cell-paragraph:" + paragraphText,
                            orderBase + blockIndex));
                    }
                } else if (child.Is(W, "tbl")) {
                    snapshots.Add(new BlockOrderSnapshot(
                        "cell-table:" + cellKey + ":" + GetTableMatchKey(partKey, child, part, options),
                        "cell-table:" + GetTableText(child),
                        orderBase + blockIndex));
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
