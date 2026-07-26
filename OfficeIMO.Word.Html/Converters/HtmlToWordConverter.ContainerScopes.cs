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

        private static List<WordParagraph> GetMaterializedContainerParagraphs(
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
            if (paragraphs.Count > 0) {
                if (breakBefore) {
                    paragraphs[0].PageBreakBefore = true;
                }
                if (breakAfter) {
                    AddPageBreakAfter(paragraphs[paragraphs.Count - 1]);
                }
                return;
            }

            if (tables.Count == 0) {
                return;
            }
            if (breakBefore) {
                InsertPageBreakAdjacentToTable(tables[0], before: true);
            }
            if (breakAfter) {
                InsertPageBreakAdjacentToTable(tables[tables.Count - 1], before: false);
            }
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
