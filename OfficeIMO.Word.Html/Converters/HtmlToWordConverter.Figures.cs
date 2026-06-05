using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ProcessFigureElement(
            IElement element,
            WordDocument doc,
            WordSection section,
            HtmlToWordOptions options,
            WordParagraph? currentParagraph,
            Stack<WordList> listStack,
            TextFormatting formatting,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordList? headingList) {
            WordParagraph? figureParagraph = currentParagraph;
            List<IElement> captions = new();

            foreach (var child in element.ChildNodes) {
                if (child is IElement childElement && string.Equals(childElement.TagName, "figcaption", StringComparison.OrdinalIgnoreCase)) {
                    captions.Add(childElement);
                    continue;
                }

                int startIndex = GetParagraphsInScope(section, cell, headerFooter).Count;
                ProcessNode(child, doc, section, options, figureParagraph, listStack, formatting, cell, headerFooter, headingList);
                if (figureParagraph == null) {
                    var paragraphs = GetParagraphsInScope(section, cell, headerFooter);
                    if (paragraphs.Count > startIndex) {
                        figureParagraph = paragraphs[paragraphs.Count - 1];
                    }
                }
            }

            foreach (var caption in captions) {
                ProcessFigureCaptionElement(caption, doc, section, options, listStack, formatting, cell, headerFooter, headingList);
            }
        }

        private void ProcessFigureCaptionElement(
            IElement caption,
            WordDocument doc,
            WordSection section,
            HtmlToWordOptions options,
            Stack<WordList> listStack,
            TextFormatting formatting,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordList? headingList) {
            ApplyCssToElement(caption);
            var paragraph = AddParagraphInScope(section, cell, headerFooter);
            paragraph.SetStyleId("Caption");
            ApplyParagraphStyleFromCss(paragraph, caption);
            ApplyClassStyle(caption, paragraph, options);
            AddBookmarkIfPresent(caption, paragraph);
            foreach (var captionChild in caption.ChildNodes) {
                ProcessNode(captionChild, doc, section, options, paragraph, listStack, formatting, cell, headerFooter, headingList);
            }
        }
    }
}
