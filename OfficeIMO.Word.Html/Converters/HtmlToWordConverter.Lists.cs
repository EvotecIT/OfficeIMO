using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System.Collections.Generic;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessList(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, TextFormatting formatting) {
            WordList list;
            if (element.TagName.Equals("ul", StringComparison.OrdinalIgnoreCase)) {
                list = cell != null ? cell.AddList(WordListStyle.Bulleted) : doc.AddListBulleted();
            } else {
                list = cell != null ? cell.AddList(WordListStyle.Headings111) : doc.AddListNumbered();
            }
            listStack.Push(list);
            foreach (var li in element.Children.OfType<IHtmlListItemElement>()) {
                ProcessListItem(li, doc, section, options, listStack, formatting, cell);
            }
            listStack.Pop();
        }

        private void ProcessListItem(IHtmlListItemElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell) {
            var list = listStack.Peek();
            int level = listStack.Count - 1;
            var paragraph = list.AddItem("", level);
            foreach (var child in element.ChildNodes) {
                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
            }
        }
    }
}