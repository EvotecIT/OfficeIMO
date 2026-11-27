using AngleSharp.Dom;
using OfficeIMO.Internal;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private WordParagraph AddParagraphInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            WordParagraph paragraph;
            if (cell != null) {
                var paragraphs = cell.Paragraphs;
                bool removeExisting = paragraphs.Count == 1 && IsReplaceableEmptyCellParagraph(paragraphs[0]);
                paragraph = cell.AddParagraph("", removeExisting);
            } else {
                paragraph = headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            }
            ApplyDefaultParagraphStyle(paragraph, cell);
            return paragraph;
        }

        private static bool IsReplaceableEmptyCellParagraph(WordParagraph paragraph) {
            foreach (OfficeXmlNode child in paragraph.Node().Children()) {
                if (IsMeaningfulCellParagraphChild(child)) return false;
            }
            return true;
        }

        private static bool IsMeaningfulCellParagraphChild(OfficeXmlNode child) {
            if (child.IsW("pPr")) {
                return child.HasChildren() || child.HasAttributes;
            }

            if (child.IsW("r")) {
                foreach (OfficeXmlNode runChild in child.Children()) {
                    if (!runChild.IsW("rPr") &&
                        (!runChild.IsW("t") || !string.IsNullOrEmpty(runChild.InnerText))) {
                        return true;
                    }
                }
                return false;
            }

            return true;
        }

        private static List<WordParagraph> GetParagraphsInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) =>
            cell?.Paragraphs ?? headerFooter?.Paragraphs ?? section.Paragraphs;

        private static List<WordTable> GetTablesInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) =>
            cell?.DirectNestedTables ?? headerFooter?.Tables ?? section.Tables;

        private List<WordParagraph> GetMaterializedContainerParagraphs(
            IElement element,
            WordSection section,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordParagraph? currentParagraph,
            ScopeMark mark) {
            List<WordParagraph> paragraphs = GetGeneratedParagraphs(section, cell, headerFooter, mark);
            List<WordTable> tables = GetGeneratedTables(section, cell, headerFooter, mark);
            paragraphs.RemoveAll(paragraph => IsGeneratedNestedTableTrailingAnchor(paragraph, tables));
            if (paragraphs.Count == 0 &&
                currentParagraph != null &&
                IsParagraphInScope(section, cell, headerFooter, mark, currentParagraph)) {
                paragraphs.Add(currentParagraph);
            }
            if (paragraphs.Count == 0 &&
                currentParagraph == null &&
                tables.Count == 0 &&
                RequiresEmptyBlockParagraph(element)) {
                AddParagraphInScope(section, cell, headerFooter);
                paragraphs = GetGeneratedParagraphs(section, cell, headerFooter, mark);
            }

            return paragraphs;
        }

        private static bool IsGeneratedNestedTableTrailingAnchor(
            WordParagraph paragraph,
            IReadOnlyList<WordTable> generatedTables) {
            // Checking each generated table's next sibling avoids PreviousSibling, which scans from the
            // first child of the parent.
            OfficeXmlNode node = paragraph.Node();
            if (!string.IsNullOrEmpty(node.InnerText) ||
                !generatedTables.Any(table => table.Node().NextSibling == node)) {
                return false;
            }

            OfficeXmlNode properties = node.WChild("pPr");
            OfficeXmlNode spacing = properties.WChild("spacing");
            return node.Children().ToList().All(element => element.IsW("pPr")) &&
                   properties.Exists && properties.Children().Count == 1 &&
                   spacing.WAttr("before") == "0" &&
                   spacing.WAttr("after") == "0" &&
                   spacing.WAttr("line") == "0";
        }

        private static bool ShouldReuseInitialWordSection(IElement element, WordDocument doc, WordSection section) {
            if (!string.Equals(element.GetAttribute("data-word-section"), "1", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!element.ClassList.Contains("word-section")) {
                return false;
            }

            if (doc.Sections.Count != 1 || !ReferenceEquals(doc.Sections[0], section)) {
                return false;
            }

            return section.Tables.Count == 0 &&
                   section.Paragraphs.All(paragraph => string.IsNullOrWhiteSpace(paragraph.Text) && !paragraph.GetRuns().Any());
        }

        private static void ApplyContainerPageBreaksFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables) {
            bool breakBefore = StyleRequestsPageBreakBefore(element);
            bool breakAfter = StyleRequestsPageBreakAfter(element);
            OfficeXmlNode first = GetGeneratedBlockBoundary(paragraphs, tables, first: true);
            OfficeXmlNode last = GetGeneratedBlockBoundary(paragraphs, tables, first: false);
            if (first.IsNone || last.IsNone) {
                return;
            }

            if (breakBefore) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => candidate.Node() == first);
                if (paragraph != null) {
                    paragraph.PageBreakBefore = true;
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => candidate.Node() == first);
                    if (table != null) {
                        InsertPageBreakAdjacentToTable(table, before: true);
                    }
                }
            }

            if (breakAfter) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => candidate.Node() == last);
                if (paragraph != null) {
                    AddPageBreakAfter(paragraph);
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => candidate.Node() == last);
                    if (table != null) {
                        InsertPageBreakAdjacentToTable(table, before: false);
                    }
                }
            }
        }

        private static OfficeXmlNode GetGeneratedBlockBoundary(
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables,
            bool first) {
            var candidates = new List<OfficeXmlNode>(paragraphs.Count + tables.Count);
            var candidateSet = new HashSet<OfficeXmlNode>();
            foreach (WordParagraph paragraph in paragraphs) {
                OfficeXmlNode node = paragraph.Node();
                if (candidateSet.Add(node)) {
                    candidates.Add(node);
                }
            }
            foreach (WordTable table in tables) {
                OfficeXmlNode node = table.Node();
                if (candidateSet.Add(node)) {
                    candidates.Add(node);
                }
            }
            if (candidates.Count == 0) {
                return default;
            }

            OfficeXmlNode parent = candidates[0].Parent;
            if (parent.IsNone) {
                return first ? candidates[0] : candidates[candidates.Count - 1];
            }

            // Generated blocks sit at the end of their container. Walking forward from the first
            // generated paragraph or table usually reaches every other candidate, which avoids
            // scanning all earlier siblings; Open XML siblings cannot be walked backwards cheaply.
            int inParent = candidates.Count(candidate => candidate.Parent == parent);
            OfficeXmlNode firstTable = tables.Count > 0 && tables[0].Node().Parent == parent ? tables[0].Node() : default;
            if (TryFindCandidateRange(candidates[0], candidateSet, inParent, out OfficeXmlNode earliest, out OfficeXmlNode latest) ||
                (firstTable.Exists && TryFindCandidateRange(firstTable, candidateSet, inParent, out earliest, out latest)) ||
                TryFindCandidateRange(parent.FirstChild, candidateSet, inParent, out earliest, out latest)) {
                return first ? earliest : latest;
            }

            return first ? candidates[0] : candidates[candidates.Count - 1];
        }

        private static bool TryFindCandidateRange(
            OfficeXmlNode start,
            HashSet<OfficeXmlNode> candidates,
            int expected,
            out OfficeXmlNode earliest,
            out OfficeXmlNode latest) {
            earliest = default;
            latest = default;
            int found = 0;
            for (OfficeXmlNode sibling = start; sibling.Exists; sibling = sibling.NextSibling) {
                if (candidates.Contains(sibling)) {
                    if (earliest.IsNone) earliest = sibling;
                    latest = sibling;
                    found++;
                }
            }

            return found == expected;
        }

        private static void InsertPageBreakAdjacentToTable(WordTable table, bool before) {
            OfficeXmlNode tableNode = table.Node();
            OfficeXmlNode paragraph = CreatePageBreakParagraph();
            if (before) {
                tableNode.Parent.InsertBefore(paragraph, tableNode);
            } else {
                tableNode.InsertAfterSelf(paragraph);
            }
        }

        /// <summary>A detached <c>w:p</c> holding one run with a page break.</summary>
        private static OfficeXmlNode CreatePageBreakParagraph() {
            OfficeXmlNode paragraph = WordConverterNodes.CreateDetachedParagraph();
            OfficeXmlNode run = paragraph.Append(paragraph.CreateElement(WordXml.W, "r"));
            OfficeXmlNode pageBreak = run.Append(run.CreateElement(WordXml.W, "br"));
            pageBreak.SetAttribute(WordXml.W, "type", "page");
            return paragraph;
        }
    }
}
