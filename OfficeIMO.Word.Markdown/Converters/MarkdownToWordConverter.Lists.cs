using Markdig.Syntax;
using OfficeIMO.Word;
using System.Linq;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessListBlock(ListBlock listBlock, WordDocument document, MarkdownToWordOptions options, int listLevel) {
            var list = listBlock.IsOrdered ? document.AddListNumbered() : document.AddListBulleted();
            foreach (ListItemBlock listItem in listBlock) {
                var firstParagraph = listItem.FirstOrDefault() as ParagraphBlock;
                if (firstParagraph != null) {
                    var listParagraph = list.AddItem(string.Empty, listLevel);
                    ProcessInline(firstParagraph.Inline, listParagraph, options, document);
                }
                foreach (var sub in listItem.Skip(1)) {
                    ProcessBlock(sub, document, options, list, listLevel + 1);
                }
            }
        }
    }
}