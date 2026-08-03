using AngleSharp.Dom;
using OfficeIMO.Html;

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
            var materialChildren = element.ChildNodes
                .Where(child => child is IElement || !string.IsNullOrWhiteSpace(child.TextContent))
                .ToList();
            int captionIndex = materialChildren.FindIndex(child =>
                child is IElement childElement &&
                string.Equals(childElement.TagName, "figcaption", StringComparison.OrdinalIgnoreCase));
            int contentBlocks = materialChildren.Count - (captionIndex >= 0 ? 1 : 0);
            bool singleContentIsImage = contentBlocks == 1 && materialChildren.Any(child =>
                child is IElement childElement &&
                string.Equals(childElement.TagName, "img", StringComparison.OrdinalIgnoreCase));
            string? captionMarker = null;
            if (captionIndex >= 0 && contentBlocks == 1 && !singleContentIsImage) {
                _figureSequence++;
                captionMarker = (captionIndex == 0 ? "officeimoFigureBefore" : "officeimoFigureAfter") +
                    _figureSequence.ToString(System.Globalization.CultureInfo.InvariantCulture);
            } else if (captionIndex >= 0 && contentBlocks > 1) {
                AddDiagnostic(
                    options,
                    "HtmlFigureStructureFlattened",
                    "A figure with multiple content blocks was imported without a reciprocal figure grouping marker.",
                    "figure",
                    lossKind: HtmlConversionLossKind.Approximation);
            }

            foreach (var child in element.ChildNodes) {
                if (child is IElement childElement && string.Equals(childElement.TagName, "figcaption", StringComparison.OrdinalIgnoreCase)) {
                    ProcessFigureCaptionElement(childElement, doc, section, options, listStack, formatting, cell, headerFooter, headingList, captionMarker);
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
            WordList? headingList,
            string? figureMarker) {
            ApplyCssToElement(caption);
            var paragraph = AddParagraphInScope(section, cell, headerFooter);
            paragraph.SetStyleId("Caption");
            ApplyParagraphStyleFromCss(paragraph, caption);
            ApplyClassStyle(caption, paragraph, options);
            AddBookmarkIfPresent(caption, paragraph);
            if (!string.IsNullOrEmpty(figureMarker)) {
                WordBookmark.AddBookmark(paragraph, figureMarker!);
            }
            foreach (var captionChild in caption.ChildNodes) {
                ProcessNode(captionChild, doc, section, options, paragraph, listStack, formatting, cell, headerFooter, headingList);
            }
        }
    }
}
