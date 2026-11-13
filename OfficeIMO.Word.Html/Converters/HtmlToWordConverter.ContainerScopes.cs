using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static WordParagraph AddParagraphInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (cell != null) {
                var paragraphs = cell.Paragraphs;
                bool removeExisting = paragraphs.Count == 1 && IsReplaceableEmptyCellParagraph(paragraphs[0]);
                return cell.AddParagraph("", removeExisting);
            }

            return headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
        }

        private static bool IsReplaceableEmptyCellParagraph(WordParagraph paragraph) =>
            !paragraph._paragraph.ChildElements.Any(IsMeaningfulCellParagraphChild);

        private static bool IsMeaningfulCellParagraphChild(OpenXmlElement child) {
            if (child is ParagraphProperties properties) {
                return properties.HasChildren || properties.HasAttributes;
            }

            if (child is Run run) {
                return run.ChildElements.Any(runChild =>
                    runChild is not RunProperties &&
                    (runChild is not Text text || !string.IsNullOrEmpty(text.Text)));
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
            if (!string.IsNullOrEmpty(paragraph._paragraph.InnerText) ||
                !generatedTables.Any(table => ReferenceEquals(table._table.NextSibling(), paragraph._paragraph))) {
                return false;
            }

            ParagraphProperties? properties = paragraph._paragraph.ParagraphProperties;
            SpacingBetweenLines? spacing = properties?.GetFirstChild<SpacingBetweenLines>();
            return paragraph._paragraph.ChildElements.All(element => element is ParagraphProperties) &&
                   properties?.ChildElements.Count == 1 &&
                   spacing?.Before?.Value == "0" &&
                   spacing.After?.Value == "0" &&
                   spacing.Line?.Value == "0";
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
            OpenXmlElement? first = GetGeneratedBlockBoundary(paragraphs, tables, first: true);
            OpenXmlElement? last = GetGeneratedBlockBoundary(paragraphs, tables, first: false);
            if (first == null || last == null) {
                return;
            }

            if (breakBefore) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => ReferenceEquals(candidate._paragraph, first));
                if (paragraph != null) {
                    paragraph.PageBreakBefore = true;
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => ReferenceEquals(candidate._table, first));
                    if (table != null) {
                        InsertPageBreakAdjacentToTable(table, before: true);
                    }
                }
            }

            if (breakAfter) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => ReferenceEquals(candidate._paragraph, last));
                if (paragraph != null) {
                    AddPageBreakAfter(paragraph);
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => ReferenceEquals(candidate._table, last));
                    if (table != null) {
                        InsertPageBreakAdjacentToTable(table, before: false);
                    }
                }
            }
        }

        private static OpenXmlElement? GetGeneratedBlockBoundary(
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables,
            bool first) {
            var candidates = new List<OpenXmlElement>(paragraphs.Count + tables.Count);
            var candidateSet = new HashSet<OpenXmlElement>(OfficeReferenceEqualityComparer<OpenXmlElement>.Instance);
            foreach (WordParagraph paragraph in paragraphs) {
                if (candidateSet.Add(paragraph._paragraph)) {
                    candidates.Add(paragraph._paragraph);
                }
            }
            foreach (WordTable table in tables) {
                if (candidateSet.Add(table._table)) {
                    candidates.Add(table._table);
                }
            }
            if (candidates.Count == 0) {
                return null;
            }

            OpenXmlElement? parent = candidates[0].Parent;
            if (parent == null) {
                return first ? candidates[0] : candidates[candidates.Count - 1];
            }

            // Generated blocks sit at the end of their container. Walking forward from the first
            // generated paragraph or table usually reaches every other candidate, which avoids
            // scanning all earlier siblings; Open XML siblings cannot be walked backwards cheaply.
            int inParent = candidates.Count(candidate => ReferenceEquals(candidate.Parent, parent));
            OpenXmlElement? firstTable = tables.Count > 0 && ReferenceEquals(tables[0]._table.Parent, parent) ? tables[0]._table : null;
            if (TryFindCandidateRange(candidates[0], candidateSet, inParent, out OpenXmlElement? earliest, out OpenXmlElement? latest) ||
                (firstTable != null && TryFindCandidateRange(firstTable, candidateSet, inParent, out earliest, out latest)) ||
                TryFindCandidateRange(parent.FirstChild, candidateSet, inParent, out earliest, out latest)) {
                return first ? earliest : latest;
            }

            return first ? candidates[0] : candidates[candidates.Count - 1];
        }

        private static bool TryFindCandidateRange(
            OpenXmlElement? start,
            HashSet<OpenXmlElement> candidates,
            int expected,
            out OpenXmlElement? earliest,
            out OpenXmlElement? latest) {
            earliest = null;
            latest = null;
            int found = 0;
            for (OpenXmlElement? sibling = start; sibling != null; sibling = sibling.NextSibling()) {
                if (candidates.Contains(sibling)) {
                    earliest ??= sibling;
                    latest = sibling;
                    found++;
                }
            }

            return found == expected;
        }

        private static void InsertPageBreakAdjacentToTable(WordTable table, bool before) {
            var paragraph = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
            if (before) {
                table._table.InsertBeforeSelf(paragraph);
            } else {
                table._table.InsertAfterSelf(paragraph);
            }
        }
    }
}
