using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private IReadOnlyList<OpenXmlElement> BuildLinkedInlineContent(
            IElement link,
            WordDocument document,
            WordSection section,
            HtmlToWordOptions options,
            Stack<WordList> listStack,
            TextFormatting formatting,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordList? headingList) {
            // An image must be added while its paragraph belongs to a document story. The
            // previous detached paragraph failed to create an image part. Keep a temporary
            // paragraph in the same scope, then copy its inline content into the hyperlink.
            // Keep the destination paragraph in place even when it is the sole empty cell
            // paragraph; AddParagraphInScope may replace that paragraph in a table cell.
            WordParagraph temporary = cell != null ? cell.AddParagraph("")
                : headerFooter != null ? headerFooter.AddParagraph("")
                : section.AddParagraph("");
            bool captured = false;
            _suppressAutoLinksDepth++;
            try {
                foreach (INode child in link.ChildNodes) {
                    ProcessLinkedInlineNode(child, document, section, options, temporary,
                        listStack, formatting, cell, headerFooter, headingList);
                }
                var inlineContent = temporary._paragraph.ChildElements
                    .Where(child => child is not ParagraphProperties)
                    .ToList();
                captured = true;
                return inlineContent;
            } finally {
                _suppressAutoLinksDepth--;
                // The copied drawing still uses the image part, so remove only the temporary
                // OpenXML paragraph. WordParagraph.Remove would also delete that image part.
                if (captured) temporary._paragraph.Remove();
                else temporary.Remove();
            }
        }

        private void ProcessLinkedInlineNode(
            INode node,
            WordDocument document,
            WordSection section,
            HtmlToWordOptions options,
            WordParagraph paragraph,
            Stack<WordList> listStack,
            TextFormatting formatting,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            WordList? headingList) {
            if (node is IElement element &&
                (_blockTags.Contains(element.TagName) || HasBlockDescendant(element))) {
                var nestedFormatting = formatting;
                ApplySpanStyles(element, ref nestedFormatting);
                foreach (INode child in element.ChildNodes) {
                    ProcessLinkedInlineNode(child, document, section, options, paragraph,
                        listStack, nestedFormatting, cell, headerFooter, headingList);
                }
                return;
            }

            ProcessNode(node, document, section, options, paragraph,
                listStack, formatting, cell, headerFooter, headingList);
        }
    }
}
