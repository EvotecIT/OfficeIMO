using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.Collections.Generic;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessList(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, TextFormatting formatting) {
            WordList list;
            bool ordered = element.TagName.Equals("ol", System.StringComparison.OrdinalIgnoreCase);
            if (ordered) {
                list = cell != null ? cell.AddList(WordListStyle.Headings111) : doc.AddListNumbered();
                var level = list.Numbering.Levels[0];
                var start = element.GetAttribute("start");
                if (!string.IsNullOrEmpty(start) && int.TryParse(start, out int startVal)) {
                    level.SetStartNumberingValue(startVal);
                }
                var type = element.GetAttribute("type");
                if (!string.IsNullOrEmpty(type)) {
                    var format = type switch {
                        "a" => NumberFormatValues.LowerLetter,
                        "A" => NumberFormatValues.UpperLetter,
                        "i" => NumberFormatValues.LowerRoman,
                        "I" => NumberFormatValues.UpperRoman,
                        _ => NumberFormatValues.Decimal,
                    };
                    level._level.NumberingFormat = new NumberingFormat { Val = format };
                }
            } else {
                list = cell != null ? cell.AddList(WordListStyle.Bulleted) : doc.AddListBulleted();
                var type = element.GetAttribute("type")?.ToLowerInvariant();
                if (!string.IsNullOrEmpty(type)) {
                    var level = list.Numbering.Levels[0];
                    switch (type) {
                        case "circle":
                            level._level.LevelText.Val = "o";
                            break;
                        case "square":
                            level._level.LevelText.Val = "■";
                            break;
                        // disc is the default, nothing to change
                    }
                }
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