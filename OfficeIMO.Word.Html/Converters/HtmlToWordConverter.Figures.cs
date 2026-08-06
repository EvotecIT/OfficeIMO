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
            int paragraphStartIndex = GetGeneratedParagraphStartIndex(section, cell, headerFooter);
            int tableStartIndex = GetTablesInScope(section, cell, headerFooter).Count;
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
            WordParagraph? captionParagraph = null;

            foreach (var child in element.ChildNodes) {
                if (child is IElement childElement && string.Equals(childElement.TagName, "figcaption", StringComparison.OrdinalIgnoreCase)) {
                    captionParagraph = ProcessFigureCaptionElement(childElement, doc, section, options, listStack, formatting, cell, headerFooter, headingList);
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

            if (captionParagraph != null && !singleContentIsImage) {
                List<WordTable> generatedTables = GetGeneratedTables(section, cell, headerFooter, tableStartIndex);
                List<WordParagraph> generatedParagraphs = GetGeneratedParagraphs(section, cell, headerFooter, paragraphStartIndex);
                generatedParagraphs.RemoveAll(paragraph =>
                    ReferenceEquals(paragraph._paragraph, captionParagraph._paragraph) ||
                    IsGeneratedNestedTableTrailingAnchor(paragraph, generatedTables));
                if (currentParagraph != null && contentBlocks > 0 &&
                    !generatedParagraphs.Any(paragraph => ReferenceEquals(paragraph._paragraph, currentParagraph._paragraph)) &&
                    GetParagraphsInScope(section, cell, headerFooter)
                        .Any(paragraph => ReferenceEquals(paragraph._paragraph, currentParagraph._paragraph))) {
                    generatedParagraphs.Add(currentParagraph);
                }

                int materializedBlocks = generatedParagraphs
                    .Select(paragraph => paragraph._paragraph)
                    .Distinct()
                    .Count() + generatedTables.Select(table => table._table).Distinct().Count();
                if (materializedBlocks == 1) {
                    _figureSequence++;
                    string captionMarker = HtmlWordRoundTripMarkers.CreateFigureCaptionBookmark(
                        captionIndex == 0,
                        _figureSequence);
                    WordBookmark.AddBookmark(captionParagraph, captionMarker);
                } else {
                    AddDiagnostic(
                        options,
                        "HtmlFigureStructureFlattened",
                        "Figure content did not materialize as exactly one reciprocal Word block, so no grouping marker was emitted.",
                        "figure",
                        lossKind: OfficeConversionLossKind.Approximation);
                }
            }
        }

        private WordParagraph ProcessFigureCaptionElement(
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
            return paragraph;
        }
    }
}
