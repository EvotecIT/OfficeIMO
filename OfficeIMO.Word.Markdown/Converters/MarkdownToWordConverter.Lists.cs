using Markdig.Extensions.TaskLists;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System.Linq;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessListBlock(ListBlock listBlock, WordDocument document, MarkdownToWordOptions options, WordList? currentList, int listLevel) {
            var list = currentList ?? (listBlock.IsOrdered ? document.AddListNumbered() : document.AddListBulleted());

            foreach (ListItemBlock listItem in listBlock) {
                var firstParagraph = listItem.FirstOrDefault() as ParagraphBlock;
                ParagraphBlock? textParagraph = null;
                if (firstParagraph != null) {
                    var listParagraph = list.AddItem(string.Empty, listLevel);
                    var firstInline = firstParagraph.Inline?.FirstChild;
                    if (firstInline is TaskList task) {
                        listParagraph.AddCheckBox(task.Checked);
                        if (task.NextSibling != null) {
                            ProcessInline(task.NextSibling, listParagraph, options, document);
                        } else {
                            textParagraph = listItem.Skip(1).OfType<ParagraphBlock>().FirstOrDefault();
                            if (textParagraph != null) {
                                ProcessInline(textParagraph.Inline, listParagraph, options, document);
                            }
                        }
                        listParagraph.Text = listParagraph.Text.TrimStart();
                    } else {
                        ProcessInline(firstParagraph.Inline, listParagraph, options, document);
                    }

                    int skip = textParagraph != null ? 2 : 1;
                    foreach (var sub in listItem.Skip(skip)) {
                        ProcessBlock(sub, document, options, list, listLevel + 1);
                    }
                } else {
                    foreach (var sub in listItem) {
                        ProcessBlock(sub, document, options, list, listLevel + 1);
                    }
                }
            }
        }
    }
}