using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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

        private static int GetGeneratedParagraphStartIndex(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            List<WordParagraph> paragraphs = GetParagraphsInScope(section, cell, headerFooter);
            return cell != null &&
                   paragraphs.Count == 1 &&
                   IsReplaceableEmptyCellParagraph(paragraphs[0])
                ? 0
                : paragraphs.Count;
        }

        private static List<WordParagraph> GetGeneratedParagraphs(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, int startIndex) =>
            GetParagraphsInScope(section, cell, headerFooter).Skip(startIndex).ToList();

        private static List<WordTable> GetTablesInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) =>
            cell?.DirectNestedTables ?? headerFooter?.Tables ?? section.Tables;

        private static List<WordTable> GetGeneratedTables(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, int startIndex) =>
            GetTablesInScope(section, cell, headerFooter).Skip(startIndex).ToList();

        private List<WordParagraph> GetMaterializedContainerParagraphs(
            IElement element,
            WordSection section,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordParagraph? currentParagraph,
            int paragraphStartIndex,
            int tableStartIndex) {
            List<WordParagraph> paragraphs = GetGeneratedParagraphs(section, cell, headerFooter, paragraphStartIndex);
            List<WordTable> tables = GetGeneratedTables(section, cell, headerFooter, tableStartIndex);
            paragraphs.RemoveAll(paragraph => IsGeneratedNestedTableTrailingAnchor(paragraph, tables));
            if (paragraphs.Count == 0 &&
                currentParagraph != null &&
                GetParagraphsInScope(section, cell, headerFooter)
                    .Any(paragraph => ReferenceEquals(paragraph._paragraph, currentParagraph._paragraph))) {
                paragraphs.Add(currentParagraph);
            }
            if (paragraphs.Count == 0 &&
                currentParagraph == null &&
                tables.Count == 0 &&
                RequiresEmptyBlockParagraph(element)) {
                AddParagraphInScope(section, cell, headerFooter);
                paragraphs = GetGeneratedParagraphs(section, cell, headerFooter, paragraphStartIndex);
            }

            return paragraphs;
        }

        private static bool IsGeneratedNestedTableTrailingAnchor(
            WordParagraph paragraph,
            IReadOnlyList<WordTable> generatedTables) {
            if (!string.IsNullOrEmpty(paragraph._paragraph.InnerText) ||
                paragraph._paragraph.PreviousSibling() is not Table previousTable ||
                !generatedTables.Any(table => ReferenceEquals(table._table, previousTable))) {
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
            foreach (WordParagraph paragraph in paragraphs) {
                if (!candidates.Any(candidate => ReferenceEquals(candidate, paragraph._paragraph))) {
                    candidates.Add(paragraph._paragraph);
                }
            }
            foreach (WordTable table in tables) {
                if (!candidates.Any(candidate => ReferenceEquals(candidate, table._table))) {
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

            IEnumerable<OpenXmlElement> siblings = first
                ? parent.ChildElements
                : parent.ChildElements.Reverse();
            return siblings.FirstOrDefault(
                       sibling => candidates.Any(candidate => ReferenceEquals(candidate, sibling))) ??
                   (first ? candidates[0] : candidates[candidates.Count - 1]);
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
