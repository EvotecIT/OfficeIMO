using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private bool TryProcessExportedHeaderFooterRegion(
            IElement element,
            WordDocument doc,
            WordSection section,
            HtmlToWordOptions options,
            TextFormatting formatting,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter) {
            if (cell != null || headerFooter != null) {
                return false;
            }

            var tagName = element.TagName.ToLowerInvariant();
            bool isHeader = string.Equals(tagName, "header", StringComparison.OrdinalIgnoreCase) && HasClass(element, "word-header");
            bool isFooter = string.Equals(tagName, "footer", StringComparison.OrdinalIgnoreCase) && HasClass(element, "word-footer");
            if (!isHeader && !isFooter) {
                return false;
            }

            var type = GetHeaderFooterType(element);
            WordHeaderFooter target = isHeader ? section.GetOrCreateHeader(type) : section.GetOrCreateFooter(type);
            RemoveEmptyHeaderFooterPlaceholders(target);

            var fmt = formatting;
            var style = element.GetAttribute("style");
            if (!string.IsNullOrWhiteSpace(style)) {
                ApplySpanStyles(element, ref fmt);
            }

            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? target.AddList(WordListStyle.Headings111) : null;
            foreach (var child in element.ChildNodes) {
                if (!string.IsNullOrWhiteSpace(style) && child is IElement childElement) {
                    var merged = MergeStyles(style, childElement.GetAttribute("style"));
                    if (!string.IsNullOrEmpty(merged)) {
                        childElement.SetAttribute("style", merged);
                    }
                }

                ProcessNode(child, doc, section, options, null, listStack, fmt, null, target, headingList);
            }

            return true;
        }

        private static bool HasClass(IElement element, string className) =>
            element.ClassList.Any(name => string.Equals(name, className, StringComparison.OrdinalIgnoreCase));

        private static HeaderFooterValues GetHeaderFooterType(IElement element) {
            var type = element.GetAttribute("data-type");
            if (string.Equals(type, "first", StringComparison.OrdinalIgnoreCase)) {
                return HeaderFooterValues.First;
            }

            if (string.Equals(type, "even", StringComparison.OrdinalIgnoreCase)) {
                return HeaderFooterValues.Even;
            }

            return HeaderFooterValues.Default;
        }

        private static void RemoveEmptyHeaderFooterPlaceholders(WordHeaderFooter headerFooter) {
            var paragraphs = headerFooter.Paragraphs;
            if (paragraphs.Count == 0 || paragraphs.Any(paragraph => !IsEmptyHeaderFooterParagraph(paragraph))) {
                return;
            }

            foreach (var paragraph in paragraphs) {
                paragraph.Remove();
            }
        }

        private static bool IsEmptyHeaderFooterParagraph(WordParagraph paragraph) =>
            string.IsNullOrWhiteSpace(paragraph.Text) &&
            !paragraph.GetRuns().Any(run => run.IsImage || run.IsStructuredDocumentTag || run.IsCheckBox || run.IsDropDownList || run.IsComboBox || run.IsDatePicker);
    }
}
